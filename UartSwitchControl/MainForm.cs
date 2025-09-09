using System;
//using System.Text;
//using System.Threading.Tasks;
using System.IO.Ports;
//using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

using System.Collections.Generic;
using System.Drawing;

namespace UartSwitchControl
{
    public partial class MainForm : Form
    {
        private SerialPort? serialPort;
        private ComboBox? comboBoxPorts;
        private Button? btnConnect;
        private Button? btnRefresh;
        private CheckBox? chkAutoRefresh;
        private Timer? portScanTimer;
        private Boolean isConnected = false;

        private Dictionary<int, Button> btnOnMap = new Dictionary<int, Button>();
        private Dictionary<int, Button> btnOffMap = new Dictionary<int, Button>();

        // 傳送記錄
        private RichTextBox? rtbLog;

        // 狀態列
        private StatusStrip? statusStrip;
        private ToolStripStatusLabel? statusConnLabel; // 連線狀態
        private ToolStripStatusLabel? statusTxLabel;   // 最後一次 TX 摘要

        public MainForm()
        {

            this.Text = "UART 8路開關控制";
            this.Width = 400;
            this.Height = 440;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            CreateUI();

            // 初始化埠清單
            RefreshPorts(preserveSelection: false);
        }

        private void CreateUI()
        {
            // COM Port 下拉
            comboBoxPorts = new ComboBox
            {
                Left = 10,
                Top = 10,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            this.Controls.Add(comboBoxPorts);

            // 連線 / 中斷
            btnConnect = new Button
            {
                Text = "連線",
                Left = 140,
                Top = 10,
                Width = 60
            };
            btnConnect.Click += BtnConnect_Click;
            this.Controls.Add(btnConnect);

            // 重新整理按鈕
            btnRefresh = new Button
            {
                Text = "重新整理",
                Left = 205,
                Top = 10,
                Width = 70
            };
            btnRefresh.Click += (s, e) => RefreshPorts(preserveSelection: true);
            this.Controls.Add(btnRefresh);

            // 自動重新整理
            chkAutoRefresh = new CheckBox
            {
                Text = "自動重新整理",
                Left = 10,
                Top = 42,
                Width = 140
            };
            chkAutoRefresh.CheckedChanged += ChkAutoRefresh_CheckedChanged;
            this.Controls.Add(chkAutoRefresh);

            // 掃描計時器
            portScanTimer = new Timer { Interval = 2000 }; // 2 秒掃描一次
            portScanTimer.Tick += (s, e) => RefreshPorts(preserveSelection: true);

            // 建 8 組按鈕
            CreateButtons();

            // 傳送記錄框
            rtbLog = new RichTextBox
            {
                Left = 240,
                Top = 70,
                Width = 120,
                Height = 280,
                ReadOnly = true,
                HideSelection = false,
                WordWrap = false
            };
            this.Controls.Add(rtbLog);

            Label lblLog = new Label
            {
                Left = 240,
                Top = 47,
                Width = 200,
                Text = "傳送記錄 (HEX)："
            };
            this.Controls.Add(lblLog);

            // 清除紀錄按鈕
            Button btnClearLog = new Button
            {
                Text = "清除紀錄",
                Left = 285,
                Top = 10,
                Width = 80
            };
            btnClearLog.Click += (s, e) =>
            {
                rtbLog.Clear();
                statusTxLabel.Text = "";
            };
            Controls.Add(btnClearLog);

            // 狀態列
            statusStrip = new StatusStrip();
            statusConnLabel = new ToolStripStatusLabel("未連線");
            statusTxLabel = new ToolStripStatusLabel("") { Spring = true, TextAlign = System.Drawing.ContentAlignment.MiddleRight };
            statusStrip.Items.Add(statusConnLabel);
            statusStrip.Items.Add(new ToolStripStatusLabel(" | "));
            statusStrip.Items.Add(statusTxLabel);
            statusStrip.Dock = DockStyle.Bottom;
            this.Controls.Add(statusStrip);
        }

        private void ChkAutoRefresh_CheckedChanged(object? sender, EventArgs e)
        {
            portScanTimer.Enabled = chkAutoRefresh.Checked;
        }

        private void RefreshPorts(bool preserveSelection)
        {
            string? selectedBefore = comboBoxPorts.SelectedItem as string;

            string[] ports = SerialPort.GetPortNames()
                                       .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                       .ToArray();

            // 避免閃爍，只有在清單不同時才重建
            bool listChanged = ports.Length != comboBoxPorts.Items.Count ||
                               ports.Where((p, i) => i >= comboBoxPorts.Items.Count || !string.Equals(p, comboBoxPorts.Items[i] as string, StringComparison.OrdinalIgnoreCase)).Any();

            if (listChanged)
            {
                comboBoxPorts.BeginUpdate();
                comboBoxPorts.Items.Clear();
                comboBoxPorts.Items.AddRange(ports);
                comboBoxPorts.EndUpdate();

                if (preserveSelection && !string.IsNullOrEmpty(selectedBefore) && ports.Contains(selectedBefore, StringComparer.OrdinalIgnoreCase))
                {
                    comboBoxPorts.SelectedItem = selectedBefore;
                }
                else
                {
                    comboBoxPorts.SelectedIndex = ports.Length > 0 ? 0 : -1;
                }
            }

            // 若目前連線中的埠已不存在，提示並中斷
            if (serialPort != null && serialPort.IsOpen && !ports.Contains(serialPort.PortName, StringComparer.OrdinalIgnoreCase))
            {
                try { serialPort.Close(); } catch { /* ignore */ }
                btnConnect.Text = "連線";
                UpdateStatusConnected(false);
                MessageBox.Show($"目前連線的埠 {serialPort.PortName} 已消失，已自動中斷連線。");
            }
        }

        private void BtnConnect_Click(object? sender, EventArgs e)
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                try { serialPort.Close(); } catch { /* ignore */ }
                btnConnect.Text = "連線";
                UpdateStatusConnected(false);
                MessageBox.Show("已中斷連線");
                return;
            }

            if (comboBoxPorts.SelectedItem == null)
            {
                MessageBox.Show("請選擇一個 COM Port");
                return;
            }

            string? selectedPort = comboBoxPorts.SelectedItem.ToString();
            serialPort = new SerialPort(selectedPort, 9600, Parity.None, 8, StopBits.One);

            try
            {
                serialPort.Open();
                btnConnect.Text = "中斷";
                UpdateStatusConnected(true);
                //MessageBox.Show("已連線到 " + selectedPort);
            }
            catch (Exception ex)
            {
                MessageBox.Show("無法連線: " + ex.Message);
            }
        }

        private void UpdateStatusConnected(bool connected)
        {
            if (connected)
            {
                statusConnLabel.Text = $"已連線到 {serialPort.PortName} (9600,8,N,1)";
                isConnected = true;
            }
            else
            {
                statusConnLabel.Text = "未連線";
                isConnected = false;
            }
        }

        private void CreateButtons()
        {
            // 8 組「開/關」按鈕，從第 1 路到第 8 路
            for (int i = 1; i <= 8; i++)
            {
                // 開
                Button btnOn = new Button
                {
                    Text = $"開 {i}",
                    Width = 100,
                    Left = 10,
                    Top = 80 + (i - 1) * 35,
                    Tag = i,
                    BackColor = Color.White
                };
                btnOn.Click += (s, e) =>
                {
                    if (((Button)s).Tag is int num)
                    {
                        SendUartCommand(num, 1);
                        if(isConnected)
                        {
                            UpdateButtonColors(num, true);
                        }
                    }
                };
                this.Controls.Add(btnOn);

                // 關
                Button btnOff = new Button
                {
                    Text = $"關 {i}",
                    Width = 100,
                    Left = 120,
                    Top = 80 + (i - 1) * 35,
                    Tag = i,
                    BackColor = Color.White
                };
                btnOff.Click += (s, e) =>
                {
                    if (((Button)s).Tag is int num)
                    {
                        SendUartCommand(num, 0);
                        if (isConnected)
                        {
                            UpdateButtonColors(num, false);
                        }
                    }
                };
                this.Controls.Add(btnOff);

                btnOnMap[i] = btnOn;
                btnOffMap[i] = btnOff;
            }
        }

        private void UpdateButtonColors(int number, bool isOn)
        {
            if (btnOnMap.ContainsKey(number) && btnOffMap.ContainsKey(number))
            {
                if (isOn)
                {
                    btnOnMap[number].BackColor = Color.LimeGreen;  // 開變綠色
                    btnOffMap[number].BackColor = Color.White;
                }
                else
                {
                    btnOffMap[number].BackColor = Color.Red;       // 關變紅色
                    btnOnMap[number].BackColor = Color.White;
                }
            }
        }

        private void SendUartCommand(int number, int status)
        {
            if (serialPort == null || !serialPort.IsOpen)
            {
                MessageBox.Show("請先連線到 COM Port！");
                return;
            }

            if (number < 1 || number > 8)
            {
                MessageBox.Show("number byte 必須介於 1~8");
                return;
            }
            if (status != 0 && status != 1)
            {
                MessageBox.Show("status byte 只能是 0 或 1");
                return;
            }

            byte header = 0xA0;
            byte numByte = (byte)number;
            byte statusByte = (byte)status;
            byte sum = (byte)(header + numByte + statusByte); // 低 8 位

            byte[] data = new byte[] { header, numByte, statusByte, sum };

            try
            {
                serialPort.Write(data, 0, data.Length);

                // 紀錄並顯示送出的指令內容（HEX）
                string hex = BytesToHex(data);
                string msg = $"{DateTime.Now:HH:mm:ss} [{number}] {(status == 1 ? "ON" : "OFF")}";//  |  {hex}";
                AppendLog(msg);

                // 狀態列顯示最後一次 TX
                statusTxLabel.Text = $"最後 TX: {hex}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("發送失敗: " + ex.Message);
            }
        }

        private static string BytesToHex(byte[] data)
        {
            // 轉為 "A0 03 01 A4" 形式
            var sb = new StringBuilder(data.Length * 3);
            for (int i = 0; i < data.Length; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append(data[i].ToString("X2"));
            }
            return sb.ToString();
        }

        private void AppendLog(string line)
        {
            if (rtbLog.TextLength > 0) rtbLog.AppendText(Environment.NewLine);
            rtbLog.AppendText(line);
            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.ScrollToCaret();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                if (serialPort != null && serialPort.IsOpen)
                    serialPort.Close();
            }
            catch { /* ignore */ }

            if (portScanTimer != null)
                portScanTimer.Dispose();

            base.OnFormClosed(e);
        }
    }
}

