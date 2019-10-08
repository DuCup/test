using System;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using Advantech.Adam;
using Apax_IO_Module_Library;
using System.Windows.Forms.DataVisualization.Charting;


namespace APAX_5017H
{
    public partial class Form_APAX_5017H : Form
    {
        // Global object
        private string APAX_INFO_NAME = "APAX";
        private string DVICE_TYPE = "5017H";

        private AdamControl m_adamCtl;
        private Apax5000Config m_aConf;

        private int m_idxID;
        private int m_ScanTime_LocalSys;
        private int m_iFailCount, m_iScanCount;
        private int m_tmpidx;
        private string[] m_szSlots;// Container of all solt device type

        private bool[] m_bChMask;
        private uint m_uiChMask = 0;
        private uint m_uiBurnoutVal = 0;
        private uint m_uiBurnoutMask = 0;
        private ushort[] m_usRanges;
        private bool m_bStartFlag = false;

        //Chart圖表定義
        private Queue<double> dataQueue = new Queue<double>(100);
        private int curValue = 0;
        private int num = 10;

        public Form_APAX_5017H() //參考APAX-5017H專案中的Program.cs檔, 得知Main函式執行後將從沒有任何參數的Form_APAX_5017H()此物件開始執行
        {
            InitializeComponent();

            InitChart();

            m_szSlots = null;
            m_iScanCount = 0;
            m_iFailCount = 0;
            m_bChMask = new bool[AdamControl.APAX_MaxAIOCh];
            m_bStartFlag = false;
            m_idxID = -1; // Set in invalid num 
            m_ScanTime_LocalSys = 500;// Scan time default 500 ms
            timer1.Interval = m_ScanTime_LocalSys;
            this.StatusBar_IO.Text = ("Start to demo "
                        + (APAX_INFO_NAME + ("-"
                        + (DVICE_TYPE + " by clicking \'Start\' button."))));
        }

        //初始化圖表
        void InitChart()
        {
            //清除chart
            this.chart1.ChartAreas.Clear();

            //定義一個 "曲線圖名稱"，並建立他的 "曲線圖區域"
            ChartArea chartArea = new ChartArea("MyChart");
            this.chart1.ChartAreas.Add(chartArea);

            //建立它的 "容器"
            this.chart1.Series.Clear();
            Series series = new Series("S1");
            series.ChartArea = "MyChart";
            this.chart1.Series.Add(series);

            //設定圖表的Y軸
            this.chart1.ChartAreas[0].AxisY.Minimum = -10;
            this.chart1.ChartAreas[0].AxisY.Maximum = 10;

            //設定圖表的x軸，時間的間隔
            this.chart1.ChartAreas[0].AxisY.Interval = 10;

            //設定圖表顏色
            this.chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.Color.Silver;
            this.chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.Color.Silver;

            //設定圖表標題
            this.chart1.Titles.Clear();
            this.chart1.Titles.Add("AI輸出");
            this.chart1.Titles[0].Text = "AI輸出顯示";
            this.chart1.Titles[0].ForeColor = Color.RoyalBlue;
            this.chart1.Titles[0].Font = new System.Drawing.Font("Microsoft Sans Serif", 12f);

            //設定曲線顏色
            this.chart1.Series[0].Color = Color.Red;

            //設定曲線圖的樣式 - 折線圖
            this.chart1.Series[0].ChartType = SeriesChartType.Line;
        }

        //更新資料佇列
        void UpdateQueue(String input)
        {
            if (dataQueue.Count > 100)
            {
                //超過100筆即刪除
                dataQueue.Dequeue();
            } 
            else
            {
                //把接收的資料寫入佇列
                dataQueue.Enqueue(Double.Parse(input));
            }

        }

        public Form_APAX_5017H(int SlotNum, int ScanTime) //因為本專案的程式開頭為Form_APAX_5017H()物件, 且本類別public Form_APAX_5017H(int SlotNum, int ScanTime)因為從頭到尾都沒有被new出來, 當成物件使用過, 因此本類別內的東西連一個都不會被執行到
        {
            InitializeComponent();
            m_szSlots = null;
            m_idxID = SlotNum; // Set Slot_ID
            m_iScanCount = 0;
            m_iFailCount = 0;
            m_bChMask = new bool[AdamControl.APAX_MaxAIOCh];
            m_bStartFlag = false;
            m_ScanTime_LocalSys = ScanTime;// Scan time
            timer1.Interval = m_ScanTime_LocalSys;
            this.StatusBar_IO.Text = ("Start to demo "
                        + (APAX_INFO_NAME + ("-"
                        + (DVICE_TYPE + " by clicking \'Start\' button."))));
        }
        
        /// <summary>
        /// Used for change I/O module 
        /// </summary>
        /// <returns></returns>
        public bool FreeResource()//將使用者於UI按下Stop後, 將會執行此函式來釋放資源
        {
            if (m_bStartFlag)
            {
                m_bStartFlag = false;
                this.tabControl1.Enabled = false;
                this.tabControl1.Visible = false;
                timer1.Enabled = false;

                m_adamCtl.Configuration().SYS_SetLocateModule(m_idxID, 0);
                m_adamCtl = null;
            }
            return true;
        }

        //開啟裝置
        public bool OpenDevice()
        {
            m_adamCtl = new AdamControl(AdamType.Apax5000);
            if (m_adamCtl.OpenDevice())
            {
                if (!m_adamCtl.Configuration().SYS_SetDspChannelFlag(false))
                {
                    this.StatusBar_IO.Text = "SYS_SetDspChannelFlag(false) Failed! ";
                    return false;
                }
                if (!m_adamCtl.Configuration().GetSlotInfo(out m_szSlots))
                {
                    this.StatusBar_IO.Text = "GetSlotInfo() Failed! ";
                    return false;
                }
            }
            return true;
        }

        //尋找裝置
        public bool DeviceFind()
        {
            int iLoop = 0;
            int iDeviceNum = 0;
            if ((m_idxID == -1))
            {
                for (iLoop = 0; iLoop < m_szSlots.Length; iLoop++)
                {
                    if ((m_szSlots[iLoop] == null))
                        continue;
                    if ((string.Compare(m_szSlots[iLoop], 0, DVICE_TYPE, 0, DVICE_TYPE.Length) == 0) && (m_szSlots[iLoop].Length == 5))
                    {
                        iDeviceNum++;
                        if ((iDeviceNum == 1))// Record first find device
                        {

                            m_idxID = iLoop;// Get DVICE_TYPE Solt

                        }
                    }
                }
            }
            else if ((string.Compare(m_szSlots[m_idxID], 0, DVICE_TYPE, 0, DVICE_TYPE.Length) == 0) && (m_szSlots[m_idxID].Length == 5))
            {
                iDeviceNum++;
            }

            if ((iDeviceNum == 1))
            {
                return true;
            }
            else if ((iDeviceNum > 1))
            {
                MessageBox.Show("Found " + iDeviceNum.ToString() + " " + DVICE_TYPE + " devices." + " It's will demo Solt " + m_idxID.ToString() + ".", "Warning");
                return true;
            }
            else
            {
                MessageBox.Show(("Can\'t find any "
                                + (DVICE_TYPE + " device!")), "Error");
                return false;
            }
        }

        //"開始"按鈕
        private void btnStart_Click(object sender, EventArgs e) //當UI上的Start按鈕被按下去後
        {
            string strbtnStatus = this.btnStart.Text;
            if ((string.Compare(strbtnStatus, "Start", true) == 0)) //進行字串比對後, 若當前UI畫面上的字串是"Start"
            {
                // Was Stop, Then Start
                if (!StartRemote())
                {
                    return;
                }
                m_bStartFlag = true;
                this.btnStart.Text = "Stop";
            }
            else  //進行字串比對後, 若當前UI畫面上的字串不是"Start"
            {
                // Was Start, Then Stop
                this.StatusBar_IO.Text = ("Start to demo "
                            + APAX_INFO_NAME + "-"
                            + DVICE_TYPE + " by clicking 'Start'button.");
                this.FreeResource(); //用於釋放資源的函式
                this.btnStart.Text = "Start";
            }
        }

        //"結束"按鈕
        private void Btn_Quit_Click(object sender, EventArgs e)  //當UI上的Quit按鈕被按下去後
        {
            Close();
        }

        //開始搖控
        public bool StartRemote()
        {
            try
            {
                if (!OpenDevice())
                {
                    throw new System.Exception("Open Local Device Fail.");
                }
                if (!DeviceFind())
                {
                    throw new System.Exception("Find " + DVICE_TYPE + "Device Fail.");
                }
                if (!RefreshConfiguration())
                {
                    throw new System.Exception("Get" + DVICE_TYPE + " Device Configuration Fail.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
                return false;
            }
            this.StatusBar_IO.Text = "Starting to remote...";
            this.timer1.Interval = m_ScanTime_LocalSys;
            this.tabControl1.Enabled = true;
            this.tabControl1.Visible = true;
            InitialDataTabPages();
            this.Text = (APAX_INFO_NAME + DVICE_TYPE);
            m_iScanCount = 0;
            this.tabControl1.SelectedIndex = 0;
            return true;
        }
        
        //更新設定
        /// <summary>
        /// APAX I/O module information init function
        /// </summary>
        /// <returns></returns>
        public bool RefreshConfiguration()

        {
            //string strModuleName; //看似重複宣告叫用, 故註解掉

            if (m_adamCtl.Configuration().GetModuleConfig(m_idxID, out m_aConf))
            {
                txtModule.Text = m_aConf.GetModuleName();       //Information-> ModuleA
                //strModuleName = m_aConf.GetModuleName();      //看似重複宣告叫用, 故註解掉
                txtID.Text = m_idxID.ToString();                //Information -> Switch ID
                txtSupportKernelFw.Text = m_aConf.wSupportFwVer.ToString("X04").Insert(2, ".");     //Information -> Support kernel Fw
                txtFwVer.Text = m_aConf.wFwVerNo.ToString("X04").Insert(2, ".");                    //Firmware version
                txtAIOFwVer.Text = m_aConf.wHwVer.ToString("X04").Insert(2, ".");   //AIO Firmware version
            }
            else
            {
                StatusBar_IO.Text = " GetModuleConfig(Error:" + m_adamCtl.Configuration().ApiLastError.ToString() + ") Failed! ";
                return false;
            }
            return true;
        }

        //初始化 頻道資訊
        /// <summary>
        /// Init Channel Information
        /// </summary>
        /// <param name="m_aConf">apax 5000 device object</param>
        /// 
        private void InitialDataTabPages()
        {
            int i = 0, idx = 0;
            byte type = (byte)_HardwareIOType.AI;   //APAX-5017H is AI module, Analog input = 3
            ListViewItem lvItem;
            string[] strRanges;
            ushort[] m_usRanges_supAI;

            for (i = 0; i < m_aConf.HwIoType.Length; i++)
            {
                if (m_aConf.HwIoType[i] == type)
                    idx = i;
            }
            m_tmpidx = idx;

            //init range combobox
            if (m_tmpidx == 0)
                m_usRanges_supAI = m_aConf.wHwIoType_0_Range;
            else if (m_tmpidx == 1)
                m_usRanges_supAI = m_aConf.wHwIoType_1_Range;
            else if (m_tmpidx == 2)
                m_usRanges_supAI = m_aConf.wHwIoType_2_Range;
            else if (m_tmpidx == 3)
                m_usRanges_supAI = m_aConf.wHwIoType_3_Range;
            else
                m_usRanges_supAI = m_aConf.wHwIoType_4_Range;
            //Get combobox items of Range
            strRanges = new string[m_aConf.HwIoType_TotalRange[m_tmpidx]];

            //顯示5種不同的電壓與電流範圍, mV和V為電壓; mA為電流
            for (i = 0; i < strRanges.Length; i++)
            {
                strRanges[i] = AnalogInput.GetRangeName(m_usRanges_supAI[i]);
            }
            SetRangeComboBox(strRanges);
            //Get combobox items of Burnout Detect Mode, Burnout代表燒毀, 耗盡體力/燈泡燒壞之意
            SetBurnoutFcnValueComboBox(new string[] { "Down Scale", "Up Scale" });
            //Get combobox items of Sampling rate (Hz/Ch)
            SetSampleRateComboBox(new string[] { "100", "1000" });
            //init channel information
            listViewChInfo.BeginUpdate(); //為了避免ListView因著數值持續改變而不斷重繪己身, 造成畫面閃爍及停頓, 故在開始前使用此函式進行UI控制, 並且在結束時使用EndUpdate(); 令UI資料刷新
            listViewChInfo.Items.Clear();

            //在顯示各Channels的數值之前, 先顯示*號替代之
            for (i = 0; i < m_aConf.HwIoTotal[m_tmpidx]; i++)
            {
                lvItem = new ListViewItem(_HardwareIOType.AI.ToString());   //type
                lvItem.SubItems.Add(i.ToString());      //Ch, UI顯示從1~11
                lvItem.SubItems.Add("*****");           //Value
                lvItem.SubItems.Add("*****");           //Ch Status 
                lvItem.SubItems.Add("*****");           //Range
                listViewChInfo.Items.Add(lvItem);
            }
            listViewChInfo.EndUpdate();
        }
        /// <summary>
        /// Get Range combobox string
        /// </summary>
        /// <param name="strRanges"></param>
        /// 

        //設定範圍的下拉式選單
        public void SetRangeComboBox(string[] strRanges) //設定Range的Combox, 故先清空原本內容及設定Count為0
        {
            cbxRange.BeginUpdate();
            cbxRange.Items.Clear(); //若combobox1已經設置了Items, 則當要清空Items時, 請使用comboBox1.Items.Clear();
            for (int i = 0; i < strRanges.Length; i++)
                cbxRange.Items.Add(strRanges[i]);

            if (cbxRange.Items.Count > 0)
                cbxRange.SelectedIndex = 0;
            cbxRange.EndUpdate();
        }
        /// <summary>
        /// Get Burnout detect mode value combobox string
        /// </summary>
        /// <param name="strRanges"></param>

        //設定燒出的下拉式選單
        public void SetBurnoutFcnValueComboBox(string[] strRanges)
        {
            cbxBurnoutValue.BeginUpdate();
            cbxBurnoutValue.Items.Clear();
            for (int i = 0; i < strRanges.Length; i++)
                cbxBurnoutValue.Items.Add(strRanges[i]);

            if (cbxBurnoutValue.Items.Count > 0)
                cbxBurnoutValue.SelectedIndex = 0;
            cbxBurnoutValue.EndUpdate();
        }
        /// <summary>
        /// Get Sampling rate value combobox string
        /// </summary>
        /// <param name="strSampleRate"></param>

        //設定採樣率的下拉式選單
        public void SetSampleRateComboBox(string[] strSampleRate)
        {
            cbxSampleRate.BeginUpdate();
            cbxSampleRate.Items.Clear();
            for (int i = 0; i < strSampleRate.Length; i++)
                cbxSampleRate.Items.Add(strSampleRate[i]);

            if (cbxSampleRate.Items.Count > 0)
                cbxSampleRate.SelectedIndex = 0;
            cbxSampleRate.EndUpdate();
        }

        //"定位"按鈕
        private void btnLocate_Click(object sender, EventArgs e)
        {
            if (btnLocate.Text == "Enable")
            {
                if (m_adamCtl.Configuration().SYS_SetLocateModule(m_idxID, 255))
                    btnLocate.Text = "Disable";
                else
                    MessageBox.Show("Locate module failed!", "Error");
            }
            else
            {
                if (m_adamCtl.Configuration().SYS_SetLocateModule(m_idxID, 0))
                    btnLocate.Text = "Enable";
                else
                    MessageBox.Show("Locate module failed!", "Error");
            }
        }

        //計時器
        /// <summary>
        /// 定期取得Channel資訊
        /// Periodically get Channel Information every time interval
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            bool bRet;
            StatusBar_IO.Text = "Polling (Interval=" + timer1.Interval.ToString() + "ms): ";
            bRet = RefreshData(); //使用此函式來刷新UI中的各個Channels的AI數值
            if (bRet) //如果上方的函式RefreshData();運行都順利且取得回傳值為true, 則執行本判斷式的內容
            {
                m_iScanCount++;
                m_iFailCount = 0;
                StatusBar_IO.Text += m_iScanCount.ToString() + " times...";
            }
            else //如果上方的函式RefreshData();運行失敗且取得回傳值為false, 則執行本判斷式的內容
            {
                m_iFailCount++;
                StatusBar_IO.Text += m_iFailCount.ToString() + " failures...";
            }

            if (m_iFailCount > 5)  //若是讀取AI數值的失敗次數超過5次, 則跳出訊息視窗顯示"輪詢暫緩(polling suspended)"等文字提醒
            {
                timer1.Enabled = false;
                StatusBar_IO.Text += " polling suspended!!";
                MessageBox.Show("Failed more than 5 times! Please check the physical connection!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1);
            }

            //by du
            //曲線圖清除
            this.chart1.Series[0].Points.Clear();
            for (int i = 0; i < dataQueue.Count; i++)
            {
                //一次印一張圖，讓曲線圖有動態的感覺
                this.chart1.Series[0].Points.AddXY((i + 1), dataQueue.ElementAt(i));
            }

            if (m_iScanCount % 50 == 0) //每執行掃描50次, 則進行資源回收動作
                GC.Collect();
        }
        /// <summary>
        /// Refresh AI Channel Information 
        /// </summary>
        /// <returns></returns>


        //※※※※更新資料※※※※
        private bool RefreshData() //本專案重點區域, 讀取AI值
        {
            int iChannelTotal = this.m_aConf.HwIoTotal[m_tmpidx];

            if (this.m_uiChMask != 0x00)
            {
                ushort[] usVal;
                Advantech.Adam.Apax5000_ChannelStatus[] aStatus;

                if (!m_adamCtl.AnalogInput().GetChannelStatus(m_idxID, iChannelTotal, out aStatus))
                {
                    StatusBar_IO.Text += "[GetChannelStatus] ApiErr:" + m_adamCtl.AnalogInput().ApiLastError.ToString() + " ";
                    return false;
                }
                if (!m_adamCtl.AnalogInput().GetValues(m_idxID, iChannelTotal, out usVal))
                {
                    StatusBar_IO.Text += "[GetValues] ApiErr:" + m_adamCtl.AnalogInput().ApiLastError.ToString() + " ";
                    return false;
                }

                string[] strVal = new string[iChannelTotal];
                string[] strStatus = new string[iChannelTotal];
                double[] dVals = new double[iChannelTotal];

                for (int i = 0; i < iChannelTotal; i++)
                {
                    if (m_aConf.wPktVer >= 0x0002)
                        dVals[i] = AnalogInput.GetScaledValueWithResolution(this.m_usRanges[i], usVal[i], m_aConf.wHwIoType_0_Resolution); //第二個參數usVal[i]將會回傳一個16 bits的值
                    else
                    {
                        if (m_aConf.GetModuleName() == "5017H")
                            dVals[i] = AnalogInput.GetScaledValueWithResolution(this.m_usRanges[i], usVal[i], 12);
                        else
                            dVals[i] = AnalogInput.GetScaledValue(this.m_usRanges[i], usVal[i]); //第一個參數this.m_usRanges[i]將會回傳一個16 bits的值
                    }

                    if (m_bChMask[i])
                    {
                        if (this.IsShowRawData)
                            strVal[i] = usVal[i].ToString("X04");
                        else
                        {
                            UpdateQueue(dVals[2].ToString(AnalogInput.GetFloatFormat(this.m_usRanges[i])));
                            strVal[i] = dVals[i].ToString(AnalogInput.GetFloatFormat(this.m_usRanges[i])); //參數this.m_usRanges[i]將會回傳一個float格式的文字字串, 推測這個就是顯示各Channels的AI數值的函式
                        }

                            strStatus[i] = aStatus[i].ToString();
                    }
                    else
                    {
                        strVal[i] = "*****";
                        strStatus[i] = "Disable";
                    }

                    //輸出值
                    listViewChInfo.Items[i].SubItems[2].Text = strVal[i].ToString();  //modify "Value" column, 此欄位中所顯示的就是AI數值
                    listViewChInfo.Items[i].SubItems[3].Text = strStatus[i].ToString();   //modify "Ch Status" column
                }
            }
            else
            {
                for (int i = 0; i < iChannelTotal; i++)
                {
                    listViewChInfo.Items[i].SubItems[2].Text = "******";        //modify "Value" column
                    listViewChInfo.Items[i].SubItems[3].Text = "******";        //modify "Ch Status" column
                }
            }
            return true;
        }
        /// <summary>
        /// When change tab, refresh status, timer, counter related informations
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        //
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strSelPageName = tabControl1.TabPages[tabControl1.SelectedIndex].Text;
            StatusBar_IO.Text = "";
            if (strSelPageName == "Module Information") //如果使用者把分頁切換至"Module Information"此分頁
            {
                m_iFailCount = 0;
                m_iScanCount = 0;
            }
            else if (strSelPageName == "AI")  //如果使用者把分頁切換至"AI"此分頁
            {
                RefreshRanges();
                RefreshAiSetting();
                if (m_aConf.GetModuleName() == "5017H")
                    RefreshBurnoutSetting(true, true);
                if (m_aConf.GetModuleName() == "5017")
                    RefreshBurnoutSetting(false, true);

                RefreshAiSampleRate();
            }

            if (tabControl1.SelectedIndex == 0)
                timer1.Enabled = false;
            else
            {
                timer1.Enabled = true;
                if (listViewChInfo.SelectedIndices.Count == 0)
                    listViewChInfo.Items[0].Selected = true;
            }
        }
        
        //更新範圍
        /// <summary>
        /// Get Channel information "Range" column
        /// </summary>
        /// <returns></returns>
        private bool RefreshRanges()
        {
            try
            {
                int iChannelTotal = this.m_aConf.HwIoTotal[m_tmpidx];
                if (m_adamCtl.Configuration().GetModuleConfig(m_idxID, out m_aConf))
                {
                    m_usRanges = m_aConf.wChRange;
                    m_uiChMask = m_aConf.dwChMask;
                    for (int i = 0; i < this.m_bChMask.Length; i++)
                    {
                        m_bChMask[i] = ((m_uiChMask & (0x01 << i)) > 0);
                    }
                    for (int i = 0; i < iChannelTotal; i++)
                    {
                        listViewChInfo.Items[i].SubItems[4].Text = AnalogInput.GetRangeName(m_usRanges[i]).ToString();
                    }
                }
                else
                    StatusBar_IO.Text += "GetModuleConfig(Error:" + m_adamCtl.Configuration().ApiLastError.ToString() + ") Failed! ";
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// Refresh Integration time 
        /// </summary>

        //更新AI設定
        private void RefreshAiSetting()
        {
            if (m_adamCtl.Configuration().GetModuleConfig(m_idxID, out m_aConf))
            {
                uint uiFcnParam;

                //Check if support SampleRate
                if (this.m_aConf.byFunType_0 == (byte)_FunctionType.Filter)
                {
                    uiFcnParam = m_aConf.dwFunParam_0;
                }
                else if (this.m_aConf.byFunType_1 == (byte)_FunctionType.Filter)
                {
                    uiFcnParam = m_aConf.dwFunParam_1;
                }
                else if (this.m_aConf.byFunType_2 == (byte)_FunctionType.Filter)
                {
                    uiFcnParam = m_aConf.dwFunParam_2;
                }
                else if (this.m_aConf.byFunType_3 == (byte)_FunctionType.Filter)
                {
                    uiFcnParam = m_aConf.dwFunParam_3;
                }
                else if (this.m_aConf.byFunType_4 == (byte)_FunctionType.Filter)
                {
                    uiFcnParam = m_aConf.dwFunParam_4;
                }
                else
                    return;
            }
            else
                StatusBar_IO.Text += "GetModuleConfig(Error:" + m_adamCtl.Configuration().ApiLastError.ToString() + ") Failed! ";
        }
        /// <summary>
        /// Refresh AI Burnout detect mode settings
        /// </summary>
        /// <param name="bUpdateBurnFun"></param>
        /// <param name="bUpdateBurnVal"></param>
        /// <returns></returns>
        private bool RefreshBurnoutSetting(bool bUpdateBurnFun, bool bUpdateBurnVal)
        {
            try
            {
                bool bRet = false;
                uint o_dwEnableMask;
                uint o_dwValue;

                ThreadStart newStart = new ThreadStart(showMsg);
                Thread waitThread = new Thread(newStart);
                waitThread.Start();

                if (bUpdateBurnFun)     //Get burnout mask value
                {
                    if (!m_adamCtl.AnalogInput().GetBurnoutFunEnable(m_idxID, out o_dwEnableMask))
                        bRet = false;
                    else
                    {
                        bRet = true;
                        m_uiBurnoutMask = o_dwEnableMask;
                    }
                    System.Threading.Thread.Sleep(1000);
                }
                if (bUpdateBurnVal)
                {
                    if (!m_adamCtl.AnalogInput().GetBurnoutValue(m_idxID, out o_dwValue))
                        bRet = false;
                    else
                    {
                        bRet = true;
                        m_uiBurnoutVal = o_dwValue;
                        if (m_uiBurnoutVal == 0x00000000)       //Update Burnout Detect Mode combobox value
                            cbxBurnoutValue.SelectedIndex = 0;
                        else
                            cbxBurnoutValue.SelectedIndex = 1;
                    }
                }
                return bRet;
            }
            catch
            {
                return false;
            }
        }
        public bool IsShowRawData
        {
            get
            {
                return chbxShowRawData.Checked;
            }
        }

        //設定AI採樣率
        private void RefreshAiSampleRate()
        {
            int idx = -1;
            uint uiRate;
            if (m_adamCtl.AnalogInput().GetSampleRate(m_idxID, out uiRate))
            {
                if (m_aConf.GetModuleName() == "5017")
                {
                    if (uiRate == 1)
                        idx = 0;
                    else if (uiRate == 10)
                        idx = 1;
                }
                else if (m_aConf.GetModuleName() == "5017H")
                {
                    if (uiRate == 100)
                        idx = 0;
                    else if (uiRate == 1000)
                        idx = 1;
                }
                else
                    idx = -2;

                if (idx >= 0)
                {
                    if (idx > cbxSampleRate.Items.Count - 1)
                        cbxSampleRate.SelectedIndex = -1;
                    else
                        cbxSampleRate.SelectedIndex = idx;
                }
                else
                    StatusBar_IO.Text += "GetSampleRate Index (Err : " + idx.ToString() + ") Failed! ";
            }
            else
                StatusBar_IO.Text += "GetSampleRate (Err : " + m_adamCtl.AnalogInput().ApiLastError.ToString() + ") Failed! ";
        }

        private void listViewChInfo_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < listViewChInfo.Items.Count; i++)
            {
                if (listViewChInfo.Items[i].Selected)
                {
                    LvChInfo_SelectedIndexChanged(i);
                    break;
                }
            }
        }
        /// <summary>
        /// When user select specific item of channel information, you should update channel range
        /// </summary>
        /// <param name="idxSel"></param>
        private void LvChInfo_SelectedIndexChanged(int idxSel) //當使用者在UI上先選擇某一個Channel後, 並點擊Range下拉式選單選擇不同的電壓/電流, 且再按下Apply, 則該Channel的Range就會立刻變更
        {
            this.cbxRange.SelectedIndex = GetChannelRangeIdx(AnalogInput.GetRangeName(m_usRanges[idxSel]));
            if ((m_usRanges[idxSel] <= (ushort)ApaxUnknown_InputRange.Btype_200To1820C && m_usRanges[idxSel] >= (ushort)ApaxUnknown_InputRange.Jtype_Neg210To1200C) ||  //0x0401~0x04C1
                (m_usRanges[idxSel] <= (ushort)ApaxUnknown_InputRange.Ni518_0To100 && m_usRanges[idxSel] >= (ushort)ApaxUnknown_InputRange.Pt100_3851_Neg200To850))     //0x0200~0x0321
            {
                this.chkBurnoutFcn.Enabled = true;
                this.btnBurnoutFcn.Enabled = true;
            }
            else
            {
                this.chkBurnoutFcn.Enabled = false;
                this.btnBurnoutFcn.Enabled = false;
            }
            //refresh burnout mask
            if (((m_uiBurnoutMask >> idxSel) & 0x1) > 0)
                chkBurnoutFcn.Checked = true;
            else
                chkBurnoutFcn.Checked = false;
        }
        public int GetChannelRangeIdx(string o_szRangeName)
        {
            for (int i = 0; i < cbxRange.Items.Count; i++)
            {
                if (cbxRange.Items[i].ToString() == o_szRangeName)
                    return i;
            }
            return -1;
        }
        private void btnApplySelRange_Click(object sender, EventArgs e)
        {
            if (!CheckControllable())
                return;
            timer1.Enabled = false;

            bool bRet = true;
            if (listViewChInfo.SelectedIndices.Count == 0 && !chkApplyAll.Checked)
            {
                MessageBox.Show("Please select the target channel in the listview!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                bRet = false;
            }
            if (bRet)
            {
                int iChannelTotal = this.m_aConf.HwIoTotal[m_tmpidx];
                ushort[] usRanges = new ushort[m_usRanges.Length];
                Array.Copy(m_usRanges, 0, usRanges, 0, m_usRanges.Length);
                if (chkApplyAll.Checked)
                {
                    for (int i = 0; i < usRanges.Length; i++)
                    {
                        usRanges[i] = AnalogInput.GetRangeCode2Byte(cbxRange.SelectedItem.ToString());
                    }
                }
                else
                {
                    for (int i = 0; i < listViewChInfo.SelectedIndices.Count; i++)
                    {
                        usRanges[listViewChInfo.SelectedIndices[i]] = AnalogInput.GetRangeCode2Byte(cbxRange.SelectedItem.ToString());
                    }
                }
                if (m_adamCtl.AnalogInput().SetRanges(this.m_idxID, iChannelTotal, usRanges))
                {
                    RefreshRanges();
                }
                else
                {
                    MessageBox.Show("Set ranges failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                }
            }
            timer1.Enabled = true;
        }
        /// <summary>
        /// Check module controllable
        /// </summary>
        /// <returns></returns>
        private bool CheckControllable()
        {
            ushort active;
            if (m_adamCtl.Configuration().SYS_GetGlobalActive(out active))
            {
                if (active == 1)
                    return true;
                else
                {
                    MessageBox.Show("There is another controller taking control, so you only can monitor IO data.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                    return false;
                }
            }
            MessageBox.Show("Checking controllable failed, utility only could monitor io data now. (" + m_adamCtl.Configuration().ApiLastError.ToString() + ")", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
            return false;
        }
        /// <summary>
        /// Waiting Foam dialog //讀取進度條視窗
        /// </summary>
        private void showMsg()
        {
            Wait_Form FormWait = new Wait_Form();
            FormWait.Start_Wait(3000); //預設為3000ms
            FormWait.ShowDialog();
            FormWait.Dispose();
            FormWait = null;
        }
        private void btnBurnoutFcn_Click(object sender, EventArgs e)
        {
            if (!CheckControllable())
                return;
            timer1.Enabled = false;
            if (chkApplyAll.Checked)
            {
                int iChannelTotal = this.m_aConf.HwIoTotal[m_tmpidx];
                if (chkBurnoutFcn.Checked)
                    m_uiBurnoutMask = (uint)(0x1 << iChannelTotal) - 1;
                else
                    m_uiBurnoutMask = 0x0;
            }
            else
            {
                int idx = 0;
                for (int i = 0; i < listViewChInfo.Items.Count; i++)
                {
                    if (listViewChInfo.Items[i].Selected)
                    {
                        idx = i;
                        break;
                    }
                }
                uint uiMask = (uint)(0x1 << idx);
                if (chkBurnoutFcn.Checked)
                    m_uiBurnoutMask |= uiMask;
                else
                    m_uiBurnoutMask &= ~uiMask;
            }
            if (m_adamCtl.AnalogInput().SetBurnoutFunEnable(m_idxID, m_uiBurnoutMask))
            {
                MessageBox.Show("Set burnout enable function done!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                RefreshBurnoutSetting(true, false); //refresh burnout mask value
            }
            else
                MessageBox.Show("Set burnout enable function failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
            timer1.Enabled = true;
        }

        private void btnBurnoutValue_Click(object sender, EventArgs e)
        {
            uint uiVal;
            if (cbxBurnoutValue.SelectedIndex == 0)
                uiVal = 0;
            else
                uiVal = 0xFFFF;
            if (!CheckControllable())
                return;
            timer1.Enabled = false;
            if (m_adamCtl.AnalogInput().SetBurnoutValue(this.m_idxID, uiVal))
            {
                MessageBox.Show("Set burnout value done!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                RefreshBurnoutSetting(false, true);     //refresh burnout detect mode
            }
            else
                MessageBox.Show("Set burnout value failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
            timer1.Enabled = true;
        }

        private void btnSampleRate_Click(object sender, EventArgs e)
        {
            int iIdx = cbxSampleRate.SelectedIndex;
            if (!CheckControllable())
                return;
            timer1.Enabled = false;

            uint uiRate;

            if (m_aConf.GetModuleName() == "5017")  //5017 module的Sample Rate範圍是1Hz(?)或10Hz(?)
            {
                if (iIdx == 0)
                    uiRate = 1;
                else
                    uiRate = 10;
            }
            else //if (m_aConf.GetModuleName() == "5017H") //5017H module的Sample Rate範圍是100Hz或1000Hz
            {
                if (iIdx == 0)
                    uiRate = 100;
                else
                    uiRate = 1000;
            }
            if (m_adamCtl.AnalogInput().SetSampleRate(this.m_idxID, uiRate))
            {
                MessageBox.Show("Set sampling rate done!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
                RefreshAiSampleRate();
            }
            else
                MessageBox.Show("Set sampling rate failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);

            timer1.Enabled = true;
        }

        private void chbxHide_CheckedChanged(object sender, EventArgs e)
        {
            panel1.Visible = !chbxHide.Checked;
        }

        private void Form_APAX_5017H_FormClosing(object sender, FormClosingEventArgs e)
        {
            FreeResource(); //用於釋放資源的函式
        }

    }
}