using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using Timer = System.Windows.Forms.Timer;
using System.Media;

namespace SpeechTopicDrawingProgram
{
    public partial class MainForm : Form
    {
        // 保存所有待抽取的题目
        private List<string> topicPool = new List<string>();
        
        // 记录已抽取的题目
        private List<string> drawnTopics = new List<string>();
        
        // 随机数生成器
        private Random random = new Random();
        
        // 定义文件路径常量
        private const string TOPIC_FILE_PATH = "topics.txt";
        
        // UI控件引用
        private Label remainingLabel;
        private TextBox resultTextBox;
        private Button drawButton;
        private Button resetButton;
        private Panel mainPanel;
        private TableLayoutPanel tableLayoutPanel;

        // 倒计时相关
        private Label timerLabel;
        private Button timerResetButton;
        private Timer countdownTimer;
        private int secondsRemaining = 180;
        private bool isTimerRunning = false;
        private string soundFilePath = "alarm.wav"; // 您的音频文件路径

        public MainForm()
        {
            InitializeComponent();
            LoadTopics();
            
            // 注册窗口大小改变事件
            this.Resize += MainForm_Resize;
            
            // 初始窗口调整
            AdjustControlSizes();
        }

        // 初始化组件
        private void InitializeComponent()
        {
            // 设置窗体属性
            this.Text = "演讲题目抽签程序";
            this.Size = new Size(1800, 1200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 250);
            this.Font = new Font("Microsoft YaHei", 10F, FontStyle.Regular);
            this.MinimumSize = new Size(800, 500);
            this.Icon = SystemIcons.Application;
            
            // 创建主布局面板
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };
            
            // 创建表格布局
            tableLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };
            
            // 设置行高比例 - 调整以突出结果区域
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 12)); // 标题
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 8));  // 计数
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 65)); // 结果 
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15)); // 按钮
            
            // 创建标题面板
            Panel titlePanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(65, 105, 225)
            };
            
            // 创建标题标签
            Label titleLabel = new Label
            {
                Text = "演讲题目抽签系统",
                Font = new Font("Microsoft YaHei", 22F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                AutoSize = false
            };
            
            // 创建信息面板
            Panel infoPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(235, 235, 240)
            };
            
            // 创建剩余题目数量标签 - 增大字体
            remainingLabel = new Label
            {
                Name = "remainingLabel",
                Text = "剩余题目: 0",
                Font = new Font("Microsoft YaHei", 16F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(60, 60, 60),
                AutoSize = false
            };
            
            // 创建结果面板
            Panel resultPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(30),
                BackColor = Color.White
            };
            
            // 创建结果容器面板 - 使用更复杂的布局来支持计时器位置
            TableLayoutPanel resultContainerPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Color.White
            };
            
            // 上部分30%放计时器区域，下部分70%放结果文本 
            resultContainerPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 20));
            resultContainerPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 80));
            
            // 计时器部分面板
            Panel timerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            
            // 使用TableLayoutPanel更精确控制计时器位置
            TableLayoutPanel timerLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            
            // 在timerLayout创建后添加
            timerLayout.ColumnStyles.Clear();
            // 调整列宽比例，确保倒计时文本有足够的右侧空间
            timerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); 
            timerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            
            // 计时器组面板 - 包含计时器和重置按钮
            Panel timerControlsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };
            
            // 创建计时器标签 - 添加到 InitializeComponent() 方法中，
            timerLabel = new Label
            {
                Name = "timerLabel",
                Text = "03:00",
                Font = new Font("Microsoft YaHei", 48F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleRight,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(65, 105, 225),
                AutoSize = false
            };
            
            // 计时器控件布局
            TableLayoutPanel timerControlsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };

            timerControlsLayout.ColumnStyles.Clear();
            timerControlsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80)); 
            timerControlsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            
            // 修改重置计时器按钮
            timerResetButton = new Button
            {
                Name = "timerResetButton",
                Text = "⟳",
                Font = new Font("Segue UI Symbol", 24F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 10, 10, 10), // 修改边距，确保按钮上下留出足够空间
                BackColor = Color.FromArgb(65, 105, 225),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false // 禁止自动调整大小
            };
            
            // 重置计时器按钮
            timerResetButton = new Button
            {
                Name = "timerResetButton",
                Text = "⟳",
                Font = new Font("Segue UI Symbol", 18F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 2, 0, 2),
                BackColor = Color.FromArgb(65, 105, 225),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            timerResetButton.FlatAppearance.BorderSize = 0;
            timerResetButton.Paint += RoundButton_Paint; // 我们将修改这个Paint处理程序
            timerResetButton.Click += TimerResetButton_Click;
            
            // 创建结果显示区域
            resultTextBox = new TextBox
            {
                Name = "resultTextBox",
                Multiline = true,
                ReadOnly = true,
                Font = new Font("Microsoft YaHei", 32F, FontStyle.Bold), 
                TextAlign = HorizontalAlignment.Center,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(50, 50, 50),
                BorderStyle = BorderStyle.None
            };
            
            // 创建按钮面板
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            
            // 创建内部按钮布局
            TableLayoutPanel buttonLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(10),
                BackColor = Color.Transparent
            };
            
            buttonLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            buttonLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            buttonLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            
            // 创建抽签按钮
            drawButton = new Button
            {
                Name = "drawButton",
                Text = "抽取题目",
                Font = new Font("Microsoft YaHei", 14F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(20, 10, 20, 10),
                BackColor = Color.FromArgb(65, 105, 225),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            drawButton.FlatAppearance.BorderSize = 0;
            drawButton.Paint += RoundButton_Paint;
            drawButton.Click += DrawButton_Click;
            
            // 创建重置按钮
            resetButton = new Button
            {
                Name = "resetButton",
                Text = "重置题库",
                Font = new Font("Microsoft YaHei", 14F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 10, 10, 10),
                BackColor = Color.FromArgb(200, 200, 205),
                ForeColor = Color.FromArgb(60, 60, 60),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            resetButton.FlatAppearance.BorderSize = 0;
            resetButton.Paint += RoundButton_Paint;
            resetButton.Click += ResetButton_Click;
            
            // 创建题目管理按钮
            Button manageButton = new Button
            {
                Name = "manageButton",
                Text = "题目管理",
                Font = new Font("Microsoft YaHei", 12F, FontStyle.Regular),
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 15, 20, 15),
                BackColor = Color.FromArgb(220, 220, 225),
                ForeColor = Color.FromArgb(60, 60, 60),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            manageButton.FlatAppearance.BorderSize = 0;
            manageButton.Paint += RoundButton_Paint;
            manageButton.Click += ManageButton_Click;
            
            // 创建倒计时定时器
            countdownTimer = new Timer();
            countdownTimer.Interval = 1000; // 1秒
            countdownTimer.Tick += CountdownTimer_Tick;
            
            // 添加控件到各自的容器
            titlePanel.Controls.Add(titleLabel);
            infoPanel.Controls.Add(remainingLabel);
            
            // 将计时器相关控件添加到布局
            timerControlsLayout.Controls.Add(timerLabel, 0, 0);
            timerControlsLayout.Controls.Add(timerResetButton, 1, 0);
            timerControlsPanel.Controls.Add(timerControlsLayout);
            
            timerLayout.Controls.Add(new Panel(), 0, 0); // 占位空白
            timerLayout.Controls.Add(timerControlsPanel, 1, 0);
            timerPanel.Controls.Add(timerLayout);
            
            // 添加计时器面板和结果文本框到结果容器
            resultContainerPanel.Controls.Add(timerPanel, 0, 0);
            resultContainerPanel.Controls.Add(resultTextBox, 0, 1);
            
            resultPanel.Controls.Add(resultContainerPanel);
            
            buttonLayoutPanel.Controls.Add(drawButton, 0, 0);
            buttonLayoutPanel.Controls.Add(resetButton, 1, 0);
            buttonLayoutPanel.Controls.Add(manageButton, 2, 0);
            buttonPanel.Controls.Add(buttonLayoutPanel);
            
            tableLayoutPanel.Controls.Add(titlePanel, 0, 0);
            tableLayoutPanel.Controls.Add(infoPanel, 0, 1);
            tableLayoutPanel.Controls.Add(resultPanel, 0, 2);
            tableLayoutPanel.Controls.Add(buttonPanel, 0, 3);
            
            mainPanel.Controls.Add(tableLayoutPanel);
            this.Controls.Add(mainPanel);
        }
        
        // 倒计时触发事件
        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            if (secondsRemaining > 0)
            {
                secondsRemaining--;
                UpdateTimerDisplay();
                
                // 当剩余时间小于等于30秒时，让计时器标签变为红色
                if (secondsRemaining <= 30)
                {
                    timerLabel.ForeColor = Color.Red;
                }
            }
            else
            {
                // 时间到
                countdownTimer.Stop();
                isTimerRunning = false;
                
                // 可选：播放提示音或显示提示
                PlayCustomSound();
            }
        }
        
        // 更新计时器显示
        private void UpdateTimerDisplay()
        {
            int minutes = secondsRemaining / 60;
            int seconds = secondsRemaining % 60;
            timerLabel.Text = $"{minutes:D2}:{seconds:D2}";
        }
        
        // 重置并开始计时
        private void StartTimer()
        {
            // 重置计时器
            secondsRemaining = 180;
            UpdateTimerDisplay();
            timerLabel.ForeColor = Color.FromArgb(65, 105, 225); // 恢复蓝色
            
            // 停止当前计时器（如果正在运行）
            if (isTimerRunning)
            {
                countdownTimer.Stop();
            }
            
            // 开始新的计时
            countdownTimer.Start();
            isTimerRunning = true;
        }
        
        // 计时器重置按钮点击事件
        private void TimerResetButton_Click(object sender, EventArgs e)
        {
            StartTimer();
        }
        
        // 窗口大小改变事件处理
        private void MainForm_Resize(object sender, EventArgs e)
        {
            AdjustControlSizes();
        }
        
        // 根据窗口大小调整控件 - 完全重写此方法以提供更好地缩放效果
        private void AdjustControlSizes()
        {
            // 根据窗口大小动态调整字体
            int windowWidth = this.ClientSize.Width;
            int windowHeight = this.ClientSize.Height;
            
            // 调整结果文本框字体大小 - 使用更合适的比例因子
            float resultFontSize = Math.Min(windowHeight / 14, windowWidth / 18);
            resultFontSize = Math.Max(24, Math.Min(64, resultFontSize)); // 提高字体大小上限
            resultTextBox.Font = new Font(resultTextBox.Font.FontFamily, resultFontSize, FontStyle.Bold);
            
            // 调整剩余题目标签字体大小
            float remainingFontSize = Math.Min(windowHeight / 28, windowWidth / 40);
            remainingFontSize = Math.Max(16, Math.Min(24, remainingFontSize));
            remainingLabel.Font = new Font(remainingLabel.Font.FontFamily, remainingFontSize, FontStyle.Bold);
            
            // 调整计时器字体大小 - 更大的字体
            float timerFontSize = Math.Min(windowHeight / 12, windowWidth / 16); // 从18和24改为12和16，使字体更大
            timerFontSize = Math.Max(38, Math.Min(75, timerFontSize)); // 从28和50改为38和75，提高最小和最大值
            timerLabel.Font = new Font(timerLabel.Font.FontFamily, timerFontSize, FontStyle.Bold);
            
            // 同时调整计时器父容器的大小
            timerLabel.AutoSize = false;
            timerLabel.Height = (int)(windowHeight / 8); // 动态调整父容器的高度
            timerLabel.Width = (int)(windowWidth / 4);  // 根据需要设置宽度
            
            // 修改timerControlsPanel的大小和位置，确保水平居中对齐
            if (timerLabel.Parent != null && timerLabel.Parent.Parent != null)
            {
                // 获取计时器按钮，调整其尺寸，防止在全屏状态下变形
                if (timerLabel.Parent is TableLayoutPanel timerControlsLayout)
                {
                    foreach (Control control in timerControlsLayout.Controls)
                    {
                        if (control == timerResetButton)
                        {
                            // 限制按钮的最大尺寸，确保全屏状态下不会过大
                            int maxButtonSize = Math.Min(windowHeight / 12, windowWidth / 20);
                            maxButtonSize = Math.Max(50, Math.Min(50, maxButtonSize));
                
                            // // 调整位置使按钮和文本垂直中心对齐
                            // control.Margin = new Padding(5, (timerLabel.Height - maxButtonSize) / 2, 5, (timerLabel.Height - maxButtonSize) / 2);
                        }
                    }
                }
            }
            
            // 调整按钮字体大小
            float buttonFontSize = Math.Min(windowHeight / 40, windowWidth / 60);
            buttonFontSize = Math.Max(12, Math.Min(20, buttonFontSize));
            drawButton.Font = new Font(drawButton.Font.FontFamily, buttonFontSize, FontStyle.Bold);
            resetButton.Font = new Font(resetButton.Font.FontFamily, buttonFontSize, FontStyle.Bold);
            
            // 调整计时器重置按钮字体大小
            float timerResetButtonFontSize = Math.Min(windowHeight / 28, windowWidth / 45);
            timerResetButtonFontSize = Math.Max(16, Math.Min(24, timerResetButtonFontSize)); // 提高最小和最大值
            timerResetButton.Font = new Font("Segue UI Symbol", timerResetButtonFontSize, FontStyle.Bold);
            
            // 调整标题字体大小
            if (tableLayoutPanel.Controls[0] is Panel titlePanel && titlePanel.Controls.Count > 0)
            {
                if (titlePanel.Controls[0] is Label titleLabel)
                {
                    float titleFontSize = Math.Min(windowHeight / 30, windowWidth / 40);
                    titleFontSize = Math.Max(18, Math.Min(32, titleFontSize));
                    titleLabel.Font = new Font(titleLabel.Font.FontFamily, titleFontSize, FontStyle.Bold);
                }
            }
            
            // 文本垂直居中的处理 - 调整TextBox文本垂直位置的处理

            resultTextBox.Multiline = true;
            
            // 设置内容垂直居中的技巧 - 调整顶部间距，适当下移位置
            int textHeight = TextRenderer.MeasureText("测试", resultTextBox.Font).Height;
            int availableHeight = resultTextBox.Height;
            
            // 修改下移量计算 - 增加额外的下移量，特别是在全屏模式下
            bool isFullscreen = (this.WindowState == FormWindowState.Maximized);
            int extraDownOffset = isFullscreen ? textHeight : 0; // 全屏时增加额外下移
            
            // 计算顶部边距，配合下移量
            int topMargin = Math.Max(0, (availableHeight - textHeight) / 2 + extraDownOffset / 2);
            
            // 通过添加空行来调整垂直位置
            if (!string.IsNullOrEmpty(resultTextBox.Text))
            {
                string originalText = resultTextBox.Text.Trim();
                int linesToAdd = Math.Max(0, topMargin / textHeight);
                
                string newText = "";
                for (int i = 0; i < linesToAdd; i++)
                {
                    newText += Environment.NewLine;
                }
                newText += originalText;
                
                // 如果文本发生变化，更新文本框
                if (resultTextBox.Text != newText)
                {
                    resultTextBox.Text = newText;
                }
            }
        }
        
        private void PlayCustomSound()
        {
            try
            {
                // 检查文件是否存在
                if (File.Exists(soundFilePath))
                {
                    using (SoundPlayer player = new SoundPlayer(soundFilePath))
                    {
                        player.Play();
                    }
                }
                else
                {
                    // 如果文件不存在，回退到系统声音
                    SystemSounds.Exclamation.Play();
            
                    // 可选：记录日志或显示消息
                    Console.WriteLine($"声音文件不存在: {soundFilePath}");
                }
            }
            catch (Exception ex)
            {
                // 出错时回退到系统声音
                SystemSounds.Exclamation.Play();
        
                // 可选：记录异常
                Console.WriteLine($"播放声音时出错: {ex.Message}");
            }
        }
        
        // 绘制圆角按钮
        private void RoundButton_Paint(object sender, PaintEventArgs e)
        {
            Button? btn = sender as Button;
            if (btn != null)
            {
                GraphicsPath path;

                // 为计时器重置按钮设置圆形
                if (btn == timerResetButton)
                {
                    // 确保按钮是正方形，取较小的尺寸
                    int size = Math.Min(btn.Width, btn.Height);
                    int offsetX = (btn.Width - size) / 2;
                    int offsetY = (btn.Height - size) / 2;
            
                    // 创建圆形路径，半径缩小50%
                    path = new GraphicsPath();
                    // 计算缩小后的大小和偏移量
                    int reducedSize = size / 2; // 减小尺寸为原来的50%
                    int newOffsetX = offsetX + (size - reducedSize) / 2;
                    int newOffsetY = offsetY + (size - reducedSize) / 2;
                    path.AddEllipse(newOffsetX, newOffsetY, reducedSize, reducedSize);
                    int maxSize = 50; // 设置最大尺寸限制
                    if (size > maxSize)
                    {
                        size = maxSize;
                    }
                }
                else
                {
                    // 其他按钮保持圆角矩形
                    int radius = 15;
                    path = new GraphicsPath();
                    path.AddArc(0, 0, radius, radius, 180, 90);
                    path.AddArc(btn.Width - radius, 0, radius, radius, 270, 90);
                    path.AddArc(btn.Width - radius, btn.Height - radius, radius, radius, 0, 90);
                    path.AddArc(0, btn.Height - radius, radius, radius, 90, 90);
                    path.CloseAllFigures();
                }

                btn.Region = new Region(path);

                // 添加轻微的阴影效果
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                // 按钮悬停时添加高亮效果
                if (btn.ClientRectangle.Contains(btn.PointToClient(Cursor.Position)))
                {
                    using (SolidBrush brush = new SolidBrush(Color.FromArgb(30, Color.White)))
                    {
                        e.Graphics.FillPath(brush, path);
                    }
                }

            }
        }

        // 加载题目数据
        private void LoadTopics()
        {
            try
            {
                if (File.Exists(TOPIC_FILE_PATH))
                {
                    string[] topics = File.ReadAllLines(TOPIC_FILE_PATH);
                    topicPool = new List<string>(topics);
                    
                    // 过滤空行
                    topicPool.RemoveAll(string.IsNullOrWhiteSpace);
                    
                    // 更新剩余题目计数
                    UpdateRemainingLabel();
                }
                else
                {
                    // 如果文件不存在,创建一个包含默认题目的文件
                    string[] defaultTopics = new string[]
                    {
                        "智慧水务在工作中的应用",
                        "新时代青年员工的责任与担当",
                        "数字化转型对企业的影响",
                        "如何提高工作效率",
                        "团队协作的重要性",
                        "创新思维与问题解决",
                        "职业规划与个人发展",
                        "环保与可持续发展",
                        "沟通技巧与人际关系",
                        "工作与生活的平衡"
                    };
                    
                    File.WriteAllLines(TOPIC_FILE_PATH, defaultTopics);
                    topicPool = new List<string>(defaultTopics);
                    UpdateRemainingLabel();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载题目失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 更新剩余题目标签
        private void UpdateRemainingLabel()
        {
            if (remainingLabel != null)
            {
                remainingLabel.Text = $"剩余题目: {topicPool.Count}";
                
                // 如果题目池为空,禁用抽取按钮
                if (drawButton != null)
                {
                    drawButton.Enabled = topicPool.Count > 0;
                    if (!drawButton.Enabled)
                    {
                        drawButton.BackColor = Color.FromArgb(180, 180, 180);
                    }
                    else
                    {
                        drawButton.BackColor = Color.FromArgb(65, 105, 225);
                    }
                }
            }
        }

        // 抽取按钮点击事件
        private void DrawButton_Click(object sender, EventArgs e)
        {
            if (topicPool.Count == 0)
            {
                MessageBox.Show("题目已抽完,请重置题库!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 随机抽取一个题目
            int index = random.Next(topicPool.Count);
            string drawnTopic = topicPool[index];
            
            // 从题目池中移除已抽取的题目
            topicPool.RemoveAt(index);
            
            // 添加到已抽取列表
            drawnTopics.Add(drawnTopic);
            
            // 显示抽取结果并增加动画效果
            ShowResultWithAnimation(drawnTopic);
            
            // 更新剩余题目计数
            UpdateRemainingLabel();
            
            // 开始或重置倒计时
            StartTimer();
        }
        
        // 结果显示动画效果 - 增强动画效果
        private async void ShowResultWithAnimation(string result)
        {
            // 保存原来的背景色
            Color originalColor = resultTextBox.BackColor;
            
            // 清空结果
            resultTextBox.Text = "";
            resultTextBox.BackColor = Color.FromArgb(240, 248, 255); // 轻微地闪烁效果
            
            // 延迟一小段时间
            await System.Threading.Tasks.Task.Delay(300);
            
            // 显示结果
            resultTextBox.Text = result;
            resultTextBox.BackColor = originalColor;
            
            // 应用居中效果
            AdjustControlSizes();
        }

        // 重置按钮点击事件
        private void ResetButton_Click(object sender, EventArgs e)
        {
            // 确认是否重置
            DialogResult result = MessageBox.Show(
                "确定要重置题库吗?\n这将恢复所有已抽取的题目。", 
                "确认", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                // 重新加载题目
                LoadTopics();
                
                // 清空结果显示
                resultTextBox.Text = "";
                
                // 清空已抽取列表
                drawnTopics.Clear();
                
                MessageBox.Show("题库已重置!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 题目管理按钮点击事件
        private void ManageButton_Click(object sender, EventArgs e)
        {
            // 打开题目管理窗口
            TopicManagementForm managementForm = new TopicManagementForm(TOPIC_FILE_PATH);
            if (managementForm.ShowDialog() == DialogResult.OK)
            {
                // 如果题目有更改,重新加载
                LoadTopics();
            }
        }
    }

    // 题目管理窗口类
    public class TopicManagementForm : Form
    {
        private string topicFilePath;
        private TextBox topicsTextBox;

        public TopicManagementForm(string filePath)
        {
            topicFilePath = filePath;
            InitializeComponent();
            LoadTopics();
            
            // 设置窗口大小跟随父窗口
            if (Owner != null)
            {
                this.Size = new Size(Owner.Width * 3 / 4, Owner.Height * 3 / 4);
            }
        }

        private void InitializeComponent()
        {
            // 设置窗体属性
            this.Text = "题目管理";
            this.Size = new Size(600, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei", 10F, FontStyle.Regular);
            this.MinimumSize = new Size(500, 400);
            this.BackColor = Color.FromArgb(245, 245, 250);
            this.Padding = new Padding(15);
            
            // 创建主布局
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Color.Transparent
            };
            
            // 增加标题栏高度，解决文字显示不完整问题
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            
            // 创建说明面板
            Panel instructionPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(65, 105, 225)
            };
            
            // 创建说明标签 - 调整垂直位置
            Label instructionLabel = new Label
            {
                Text = "题目管理 (每行输入一个题目)",
                Font = new Font("Microsoft YaHei", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft, // 确保文本垂直居中
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                Padding = new Padding(15, 0, 0, 0),
                AutoSize = false // 禁止自动调整大小
            };
            
            // 创建题目编辑面板
            Panel editPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
                BackColor = Color.White
            };
            
            // 创建题目编辑框
            topicsTextBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.None,
                Font = new Font("Microsoft YaHei", 12F, FontStyle.Regular),
                BackColor = Color.White
            };
            
            // 创建按钮面板
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            
            TableLayoutPanel buttonLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(10)
            };
            
            buttonLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttonLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            
            // 创建保存按钮
            Button saveButton = new Button
            {
                Text = "保存",
                Size = new Size(150, 40),
                Font = new Font("Microsoft YaHei", 12F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(50, 5, 10, 5),
                BackColor = Color.FromArgb(65, 105, 225),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            saveButton.FlatAppearance.BorderSize = 0;
            saveButton.Click += SaveButton_Click;
            
            // 创建取消按钮
            Button cancelButton = new Button
            {
                Text = "取消",
                Size = new Size(150, 40),
                Font = new Font("Microsoft YaHei", 12F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 5, 50, 5),
                BackColor = Color.FromArgb(200, 200, 205),
                ForeColor = Color.FromArgb(60, 60, 60),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            cancelButton.FlatAppearance.BorderSize = 0;
            cancelButton.Click += (s, e) => this.DialogResult = DialogResult.Cancel;
            
            // 添加控件到容器
            instructionPanel.Controls.Add(instructionLabel);
            editPanel.Controls.Add(topicsTextBox);
            
            buttonLayout.Controls.Add(saveButton, 0, 0);
            buttonLayout.Controls.Add(cancelButton, 1, 0);
            buttonPanel.Controls.Add(buttonLayout);
            
            mainLayout.Controls.Add(instructionPanel, 0, 0);
            mainLayout.Controls.Add(editPanel, 0, 1);
            mainLayout.Controls.Add(buttonPanel, 0, 2);
            
            this.Controls.Add(mainLayout);
        }

        // 加载题目
        private void LoadTopics()
        {
            try
            {
                if (File.Exists(topicFilePath))
                {
                    string content = File.ReadAllText(topicFilePath);
                    topicsTextBox.Text = content;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载题目失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 保存按钮点击事件
        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取编辑后的题目
                string content = topicsTextBox.Text;
                
                // 检查是否有内容
                if (string.IsNullOrWhiteSpace(content))
                {
                    DialogResult result = MessageBox.Show(
                        "题目列表为空，确定要保存吗？", 
                        "确认", 
                        MessageBoxButtons.YesNo, 
                        MessageBoxIcon.Question);
                        
                    if (result != DialogResult.Yes)
                    {
                        return;
                    }
                }
                
                // 保存到文件
                File.WriteAllText(topicFilePath, content);
                
                MessageBox.Show("题目保存成功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存题目失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // 程序入口点
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}