using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace reborncam
{
 public partial class Form1 : Form
 {
 private System.Windows.Forms.Timer _thumbnailTimer;
 private PictureBox[] _previewBoxes;
 private PictureBox[] _thumbnailBoxes;

 // OpenCV 関連
 private const int MaxCameraCount =6;
 private VideoCapture[] _captures = new VideoCapture[MaxCameraCount];
 private Mat[] _latestFrames = new Mat[MaxCameraCount];
 private Bitmap[] _dispBitmaps = new Bitmap[MaxCameraCount];
 private bool[] _cameraAlive = new bool[MaxCameraCount];
 private Label[] _statusLabels;

 private bool _cameraOpenFlag = false;
 private BackgroundWorker _previewWorker;
 private CancellationTokenSource _cameraCts;
 private Task<bool>[] _frameTasks;

 private string _logFilePath;
 private string _diagnosticFilePath;
 private readonly object _diagLock = new object();
 private string _currentSerialNumber = string.Empty;

 // カメラポート割り当て（PictureBox のインデックス → カメラデバイスインデックス）
 private int[] _cameraPortAssignments = new int[MaxCameraCount] {0,1,2,3,4,5 };

 // 各カメラの部位名（変更可能）
 private string[] _cameraPartNames = new[]
 {
 "Front",
 "Back",
 "Left",
 "Right",
 "Top",
 "Bottom"
 };

 // 部位名の選択肢
 private readonly string[] _availablePartNames = new[]
 {
 "Front",
 "Back",
 "Left",
 "Right",
 "Top",
 "Bottom",
 "Custom1",
 "Custom2",
 "Custom3"
 };

 // 解像度の優先順位（4K → 段階的にフォールバック）
 private readonly int[,] _resolutionPresets = new int[,]
 {
 {3840,2160 }, //4K
 {2592,1944 },
 {1920,1080 }, // FullHD
 {1280,720 },
 {640,480 },
 {320,240 }
 };

 // 設定UI
 private ComboBox _portComboBox;
 private ComboBox _partNameComboBox;
 private int _selectedCameraIndex =0;

 public Form1()
 {
 InitializeComponent();

 // ログフォルダとログファイルを初期化
 var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log");
 Directory.CreateDirectory(logDir);
 _logFilePath = Path.Combine(logDir, "app.log");
 WriteLog("Application started");

 // LOFF diagnostics folder
 var diagDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOFF");
 Directory.CreateDirectory(diagDir);
 _diagnosticFilePath = Path.Combine(diagDir, "diagnog.log");
 // write resumed marker
 WriteDiagnostic("--- diagnostics resumed " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));

 _previewBoxes = new[] { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };
 _statusLabels = new[] { statusLabelMonitor1, statusLabelMonitor2, statusLabelMonitor3, statusLabelMonitor4, statusLabelMonitor5, statusLabelMonitor6 };

 _thumbnailBoxes = new PictureBox[MaxCameraCount];
 for (int i =0; i < MaxCameraCount; i++)
 {
 var thumb = new PictureBox
 {
 SizeMode = PictureBoxSizeMode.Zoom,
 BorderStyle = BorderStyle.FixedSingle,
 BackColor = Color.Black,
 Width =120,
 Height =90,
 Margin = new Padding(2)
 };
 _thumbnailBoxes[i] = thumb;
 }

 var thumbsPanel = new FlowLayoutPanel
 {
 Dock = DockStyle.Bottom,
 Height =100,
 FlowDirection = FlowDirection.RightToLeft,
 BackColor = Color.Black
 };

 for (int i =0; i < MaxCameraCount; i++)
 {
 thumbsPanel.Controls.Add(_thumbnailBoxes[i]);
 }

 Controls.Add(thumbsPanel);

 _thumbnailTimer = new System.Windows.Forms.Timer();
 _thumbnailTimer.Interval =10_000;
 _thumbnailTimer.Tick += ThumbnailTimer_Tick;

 // BackgroundWorkerでプレビュー更新（Pileave スタイル）
 _previewWorker = new BackgroundWorker();
 _previewWorker.WorkerSupportsCancellation = true;
 _previewWorker.WorkerReportsProgress = true;
 _previewWorker.DoWork += PreviewWorker_DoWork;
 _previewWorker.ProgressChanged += PreviewWorker_ProgressChanged;
 _previewWorker.RunWorkerCompleted += PreviewWorker_Completed;

 simulationButton.Enabled = false;

 SetupCameraConfigUI();
 }

 private void SetupCameraConfigUI()
 {
 cameraSelectionComboBox.Items.Clear();
 for (int i =0; i < MaxCameraCount; i++)
 {
 cameraSelectionComboBox.Items.Add($"Camera {i +1}");
 }

 var portLabel = new Label
 {
 Text = "Camera Port:",
 AutoSize = true,
 Location = new System.Drawing.Point(15,210)
 };
 settingsPanel.Controls.Add(portLabel);

 _portComboBox = new ComboBox
 {
 DropDownStyle = ComboBoxStyle.DropDownList,
 Location = new System.Drawing.Point(3,228),
 Size = new System.Drawing.Size(154,23)
 };
 for (int i =0; i <=9; i++)
 {
 _portComboBox.Items.Add($"Port {i}");
 }
 _portComboBox.SelectedIndexChanged += PortComboBox_SelectedIndexChanged;
 settingsPanel.Controls.Add(_portComboBox);

 var partLabel = new Label
 {
 Text = "Part Name:",
 AutoSize = true,
 Location = new System.Drawing.Point(15,260)
 };
 settingsPanel.Controls.Add(partLabel);

 _partNameComboBox = new ComboBox
 {
 DropDownStyle = ComboBoxStyle.DropDownList,
 Location = new System.Drawing.Point(3,278),
 Size = new System.Drawing.Size(154,23)
 };
 _partNameComboBox.Items.AddRange(_availablePartNames);
 _partNameComboBox.SelectedIndexChanged += PartNameComboBox_SelectedIndexChanged;
 settingsPanel.Controls.Add(_partNameComboBox);

 UpdateConfigUI();

 // Now it's safe to set SelectedIndex so SelectedIndexChanged won't run before controls exist
 if (cameraSelectionComboBox.Items.Count >0)
 cameraSelectionComboBox.SelectedIndex =0;
 }

 private void UpdateConfigUI()
 {
 // guard against calling before UI controls exist
 if (_portComboBox != null && _partNameComboBox != null)
 {
 _portComboBox.SelectedIndex = Math.Clamp(_cameraPortAssignments[_selectedCameraIndex],0, _portComboBox.Items.Count -1);
 var currentPart = _cameraPartNames[_selectedCameraIndex];
 var idx = Array.IndexOf(_availablePartNames, currentPart);
 _partNameComboBox.SelectedIndex = idx >=0 ? idx :0;
 }
 }

 private async void Form1_Load_1(object? sender, EventArgs e)
 {
 WriteLog("Form1_Load: Initializing cameras...");
 statusLabel.Text = "Opening cameras...";

 _cameraOpenFlag = true;

 // 初期化画面を表示
 for (int i =0; i < MaxCameraCount; i++)
 {
 _statusLabels[i].Text = $"Camera{i +1}\nOpening...";
 _statusLabels[i].BackColor = Color.Gray;
 }

 // カメラを並列で高速オープン
 var openTasks = new Task<bool>[MaxCameraCount];
 for (int i =0; i < MaxCameraCount; i++)
 {
 int index = i;
 openTasks[i] = Task.Run(() => OpenCameraDevice(index));
 }

 var results = await Task.WhenAll(openTasks);

 // カメラ状態を UI に反映
 for (int i =0; i < MaxCameraCount; i++)
 {
 _cameraAlive[i] = results[i];
 UpdateCameraStatus(i);
 }

 WriteLog("All cameras initialized");

 // BackgroundWorkerでプレビュー開始
 _previewWorker.RunWorkerAsync();

 // フレーム取得タスクを並列起動（Pileave スタイル）
 _cameraCts = new CancellationTokenSource();
 _frameTasks = new Task<bool>[MaxCameraCount];
 for (int i =0; i < MaxCameraCount; i++)
 {
 int index = i;
 var token = _cameraCts.Token;
 _frameTasks[i] = Task.Run(() => GetCameraFrameLoop(index, token), token);
 }

 statusLabel.Text = "Previewing - Waiting for Serial Number...";
 WriteLog("Preview started");

 // バックグラウンドで監視（アプリ終了時に Task を待つため）
 _ = Task.Run(async () =>
 {
 try
 {
 await Task.WhenAll(_frameTasks.Where(t => t != null));
 }
 catch { }
 finally
 {
 if (_frameTasks != null)
 {
 foreach (var t in _frameTasks) t?.Dispose();
 _frameTasks = null;
 }
 }
 });
 }

 private bool OpenCameraDevice(int cameraIndex)
 {
 var port = _cameraPortAssignments[cameraIndex];

 try
 {
 WriteLog($"Camera {cameraIndex +1}: Attempting to open port {port}...");

 var cap = new VideoCapture();
 cap.Open(port, VideoCaptureAPIs.MSMF); // MSMFで高速化

 if (!cap.IsOpened())
 {
 WriteLog($"Camera {cameraIndex +1} (Port {port}): Failed to open");
 WriteDiagnostic(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $" | カメラ {cameraIndex +1} (ポート {port}) を開けませんでした。");
 cap.Dispose();
 return false;
 }

 // 解像度を段階的に試す（4K優先）
 bool resolutionSet = false;
 for (int r =0; r < _resolutionPresets.GetLength(0); r++)
 {
 int width = _resolutionPresets[r,0];
 int height = _resolutionPresets[r,1];

 cap.Set(VideoCaptureProperties.FrameWidth, width);
 cap.Set(VideoCaptureProperties.FrameHeight, height);

 int actualWidth = (int)cap.Get(VideoCaptureProperties.FrameWidth);
 int actualHeight = (int)cap.Get(VideoCaptureProperties.FrameHeight);

 if (actualWidth == width && actualHeight == height)
 {
 WriteLog($"Camera {cameraIndex +1} (Port {port}): Resolution set to {width}x{height}");
 resolutionSet = true;
 break;
 }
 }

 if (!resolutionSet)
 {
 var w = (int)cap.Get(VideoCaptureProperties.FrameWidth);
 var h = (int)cap.Get(VideoCaptureProperties.FrameHeight);
 WriteLog($"Camera {cameraIndex +1} (Port {port}): Using default resolution {w}x{h}");
 }

 cap.Set(VideoCaptureProperties.Fps,30);

 _captures[cameraIndex] = cap;
 _latestFrames[cameraIndex] = new Mat();

 WriteLog($"Camera {cameraIndex +1} (Port {port}): Opened successfully");
 return true;
 }
 catch (Exception ex)
 {
 WriteLog($"Camera {cameraIndex +1} (Port {port}): Exception - {ex.Message}");
 WriteDiagnostic(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $" | Camera {cameraIndex +1} FAILURE | {ex.Message}");
 return false;
 }
 }

 private bool GetCameraFrameLoop(int cameraIndex, CancellationToken token)
 {
 if (!_cameraAlive[cameraIndex])
 {
 return false;
 }

 var cap = _captures[cameraIndex];
 var frame = _latestFrames[cameraIndex];

 while (true)
 {
 try
 {
 if (!_cameraOpenFlag || this.IsDisposed || (token != null && token.IsCancellationRequested))
 {
 break;
 }

 if (cap == null || !cap.IsOpened())
 {
 // カメラが死んだ
 _cameraAlive[cameraIndex] = false;
 WriteLog($"Camera {cameraIndex +1}: Device lost during preview");
 WriteDiagnostic(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $" | Camera {cameraIndex +1} FAILURE | capture=null");
 break;
 }

 bool readSuccess = false;
 try
 {
 readSuccess = cap.Read(frame);
 }
 catch (ObjectDisposedException)
 {
 // capture was disposed from shutdown; treat as dead and exit
 _cameraAlive[cameraIndex] = false;
 WriteLog($"Camera {cameraIndex +1}: capture disposed during read");
 break;
 }

 if (!readSuccess || frame.Empty())
 {
 // 読み込み失敗が続く場合は死んだと判定
 _cameraAlive[cameraIndex] = false;
 WriteLog($"Camera {cameraIndex +1}: Frame read failed, marking as dead");
 WriteDiagnostic(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $" | Camera {cameraIndex +1} FAILURE | capture=null");
 break;
 }

 // Cv2.WaitKeyで停止回避
 Cv2.WaitKey(1);
 }
 catch (Exception ex)
 {
 _cameraAlive[cameraIndex] = false;
 WriteLog($"Camera {cameraIndex +1}: Exception in frame loop - {ex.Message}");
 break;
 }
 }

 //ループ終了時にリソース解放
 // don't dispose here if cancellation is not requested; disposal will be handled on shutdown
 try { cap?.Dispose(); } catch { }
 return true;
 }

 private void PreviewWorker_DoWork(object? sender, DoWorkEventArgs e)
 {
 while (true)
 {
 if (_previewWorker.CancellationPending || !_cameraOpenFlag)
 {
 e.Cancel = true;
 break;
 }

 Thread.Sleep(33); // 約30fps

 _previewWorker.ReportProgress(1);
 }
 }

 private void PreviewWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
 {
 // 古い画像を保持して後で破棄
 Image[] oldImages = new Image[MaxCameraCount];
 for (int i =0; i < MaxCameraCount; i++)
 {
 oldImages[i] = _previewBoxes[i].Image;
 }

 // 最新フレームを Bitmap に変換して表示
 for (int i =0; i < MaxCameraCount; i++)
 {
 if (!_cameraAlive[i])
 {
 _dispBitmaps[i] = null;
 UpdateCameraStatus(i);
 continue;
 }

 var frame = _latestFrames[i];
 if (frame != null && !frame.Empty())
 {
 try
 {
 var oldBmp = _dispBitmaps[i];
 _dispBitmaps[i] = BitmapConverter.ToBitmap(frame);
 oldBmp?.Dispose();
 }
 catch
 {
 _dispBitmaps[i] = null;
 }
 }
 }

 // UI に反映
 for (int i =0; i < MaxCameraCount; i++)
 {
 if (_dispBitmaps[i] != null)
 {
 _previewBoxes[i].Image = (Bitmap)_dispBitmaps[i].Clone();
 }
 else
 {
 _previewBoxes[i].Image = null;
 }
 }

 // 古い画像を破棄
 for (int i =0; i < oldImages.Length; i++)
 {
 oldImages[i]?.Dispose();
 }
 }

 private void PreviewWorker_Completed(object? sender, RunWorkerCompletedEventArgs e)
 {
 WriteLog("Preview worker completed");
 for (int i =0; i < MaxCameraCount; i++)
 {
 _previewBoxes[i].Image = null;
 }
 }

 private void UpdateCameraStatus(int cameraIndex)
 {
 if (_cameraAlive[cameraIndex])
 {
 _statusLabels[cameraIndex].Text = $"Camera{cameraIndex +1}\n{_cameraPartNames[cameraIndex]}\nOK";
 _statusLabels[cameraIndex].BackColor = Color.Green;
 }
 else
 {
 _statusLabels[cameraIndex].Text = $"Camera{cameraIndex +1}\n{_cameraPartNames[cameraIndex]}\nDEAD";
 _statusLabels[cameraIndex].BackColor = Color.Red;
 }
 }

 private void SerialNumberTextBox_TextChanged(object? sender, EventArgs e)
 {
 var serial = serialNumberTextBox.Text.Trim();
 if (!string.IsNullOrWhiteSpace(serial))
 {
 _currentSerialNumber = serial;
 simulationButton.Enabled = true;
 WriteLog($"Serial number entered: {serial}");
 }
 else
 {
 _currentSerialNumber = string.Empty;
 simulationButton.Enabled = false;
 }
 }

 private void SimulationButton_Click(object? sender, EventArgs e)
 {
 WriteLog($"Capture button clicked (Serial: {_currentSerialNumber})");
 CaptureCurrentFrames();
 }

 private void PreviewButton_Click(object? sender, EventArgs e)
 {
 WriteLog("Preview button clicked");
 ClearThumbnails();
 }

 private void CameraSelectionComboBox_SelectedIndexChanged(object? sender, EventArgs e)
 {
 _selectedCameraIndex = cameraSelectionComboBox.SelectedIndex;
 if (_selectedCameraIndex >=0 && _selectedCameraIndex < MaxCameraCount)
 {
 UpdateConfigUI();
 WriteLog($"Selected camera {_selectedCameraIndex +1} for configuration");
 }
 }

 private void PortComboBox_SelectedIndexChanged(object? sender, EventArgs e)
 {
 var newPort = _portComboBox.SelectedIndex;
 if (_selectedCameraIndex >=0 && _selectedCameraIndex < MaxCameraCount)
 {
 var oldPort = _cameraPortAssignments[_selectedCameraIndex];
 _cameraPortAssignments[_selectedCameraIndex] = newPort;
 WriteLog($"Camera {_selectedCameraIndex +1}: Port changed from {oldPort} to {newPort} (requires restart)");
 }
 }

 private void PartNameComboBox_SelectedIndexChanged(object? sender, EventArgs e)
 {
 var newPart = _partNameComboBox.SelectedItem?.ToString();
 if (!string.IsNullOrEmpty(newPart) && _selectedCameraIndex >=0 && _selectedCameraIndex < MaxCameraCount)
 {
 var oldPart = _cameraPartNames[_selectedCameraIndex];
 _cameraPartNames[_selectedCameraIndex] = newPart;
 WriteLog($"Camera {_selectedCameraIndex +1}: Part name changed from {oldPart} to {newPart}");
 UpdateCameraStatus(_selectedCameraIndex);
 }
 }

 private void BrightnessTrackBar_Scroll(object? sender, EventArgs e)
 {
 // TODO: 選択中のカメラの明るさ制御
 }

 private void ContrastTrackBar_Scroll(object? sender, EventArgs e)
 {
 // TODO: 選択中のカメラのコントラスト制御
 }

 private void CaptureCurrentFrames()
 {
 var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
 var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log");
 var tempDir = Path.Combine(logDir, timestamp);
 var zipPath = Path.Combine(logDir, $"{_currentSerialNumber}_{timestamp}.zip");

 try
 {
 Directory.CreateDirectory(tempDir);
 WriteLog($"Capture started: Serial={_currentSerialNumber}, Saving to {_currentSerialNumber}_{timestamp}.zip");

 int savedCount =0;
 for (int i =0; i < _previewBoxes.Length; i++)
 {
 var src = _previewBoxes[i].Image;
 if (src != null && _cameraAlive[i])
 {
 _thumbnailBoxes[i].Image?.Dispose();
 _thumbnailBoxes[i].Image = new Bitmap(src);

 var partName = _cameraPartNames[i];
 var fileName = $"{_currentSerialNumber}_{partName}_{timestamp}.jpg";
 var filePath = Path.Combine(tempDir, fileName);
 src.Save(filePath, ImageFormat.Jpeg);
 savedCount++;
 WriteLog($" Camera {i +1} (Port {_cameraPortAssignments[i]}, {partName}): Saved as {fileName}");
 }
 else if (!_cameraAlive[i])
 {
 WriteLog($" Camera {i +1}: Skipped (camera dead)");
 }
 }

 if (savedCount >0)
 {
 ZipFile.CreateFromDirectory(tempDir, zipPath, CompressionLevel.Fastest, false);
 WriteLog($"Capture completed: {savedCount} images saved to {_currentSerialNumber}_{timestamp}.zip");
 statusLabel.Text = $"Captured {savedCount} images";
 }
 else
 {
 WriteLog("Capture completed: No images saved (all cameras dead)");
 statusLabel.Text = "Capture failed - No cameras available";
 }

 Directory.Delete(tempDir, true);

 serialNumberTextBox.Clear();
 simulationButton.Enabled = false;
 WriteLog("Ready for next serial number");
 }
 catch (Exception ex)
 {
 WriteLog($"Capture error: {ex.Message}");
 MessageBox.Show($"撮影エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
 }

 _thumbnailTimer.Stop();
 _thumbnailTimer.Start();
 }

 private void ThumbnailTimer_Tick(object? sender, EventArgs e)
 {
 _thumbnailTimer.Stop();
 ClearThumbnails();
 }

 private void ClearThumbnails()
 {
 for (int i =0; i < _thumbnailBoxes.Length; i++)
 {
 _thumbnailBoxes[i].Image?.Dispose();
 _thumbnailBoxes[i].Image = null;
 }
 }

 protected override void OnFormClosing(FormClosingEventArgs e)
 {
 base.OnFormClosing(e);

 WriteLog("Application closing: Releasing cameras...");

 // signal cancellation to frame tasks and preview worker
 _cameraOpenFlag = false;
 _previewWorker.CancelAsync();
 _cameraCts?.Cancel();

 // wait for frame tasks to finish (up to2 seconds)
 try
 {
 if (_frameTasks != null)
 {
 Task.WaitAll(_frameTasks.Where(t => t != null).ToArray(),2000);
 }
 }
 catch { }

 // now safe to dispose resources
 for (int i =0; i < MaxCameraCount; i++)
 {
 try { _previewBoxes[i].Image?.Dispose(); _previewBoxes[i].Image = null; } catch { }
 try { _thumbnailBoxes[i].Image?.Dispose(); _thumbnailBoxes[i].Image = null; } catch { }
 try { _dispBitmaps[i]?.Dispose(); _dispBitmaps[i] = null; } catch { }
 try { _latestFrames[i]?.Dispose(); _latestFrames[i] = null!; } catch { }
 try { _captures[i]?.Release(); _captures[i]?.Dispose(); _captures[i] = null!; } catch { }
 }

 _cameraCts?.Dispose();
 _cameraCts = null;

 WriteLog("Application closed");
 }

 private void WriteLog(string message)
 {
 try
 {
 var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {message}";
 File.AppendAllText(_logFilePath, line + Environment.NewLine);
 Debug.WriteLine(line);
 }
 catch
 {
 // ログ書き込みエラーは無視
 }
 }

 /// <summary>
 /// Diagnostic log written to LOFF\diagnog.log (thread-safe)
 /// </summary>
 private void WriteDiagnostic(string message)
 {
 try
 {
 lock (_diagLock)
 {
 File.AppendAllText(_diagnosticFilePath, message + Environment.NewLine);
 }
 Debug.WriteLine("DIAG: " + message);
 }
 catch
 {
 // ignore
 }
 }
 }
}
