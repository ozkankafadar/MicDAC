using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms.DataVisualization.Charting;

namespace MicDAC
{
    public partial class MainForm : Form
    {
        #region Variables
        SerialPort sPort;
        Thread mainThread;
        StreamWriter sw, swrite;
        Graphics monitoringGraphics;
        Bitmap RealTimeDrawArea;
        Pen pen1,pen2,pen3;
        string data;
        string[] dataArray;
        double[] signal1, signal2, signal3;
        double[] signalReal1, signalReal2, signalReal3;
        const int N = 1024;
        double offset1,offset2,offset3;
        float sampling;
        double point1, point2;
        PointF p1, p2;
        Font font;
        decimal stopTime;
        int dataCount;
        double[,] orginalSignal,orginalSignalRev,orginalSignalRevCopy;
        public static double[,] filteredSignal;
        double cn=0.1;
        Graphics orginalSignal_Graphics,filteredSignal_Graphics;
        Bitmap orginalSignal_Bitmap,filteredSignal_Bitmap,windowing_Bitmap;
        string activeDirectory;
        bool filter = false;
        double frequency = 0;
        int window = 0;
        double bandwidth = 0;
        string filename;
        bool[] window_array;
        float[,] koord_window;
        float[,] koord_window2;
        float cn2 = 0;
        float refer = 0;
        bool isCont = false;
        string fileNameWithoutExtension, directoryName, fullName;
        double[] f1;
        public double[,] FFTnew_signalNS;
        public double[,] FFTnew_signalEW;
        public double[,] FFTnew_signalUD;
        public double[,] smoothedFFTnew_signalNS;
        public double[,] smoothedFFTnew_signalEW;
        public double[,] smoothedFFTnew_signalUD;
        public double[,] HV;
        public double[] HV_Avg;
        public double[] HV_Std;//?
        public double[] HV_StdMult;//?
        public double[] HV_StdDiv;//?
        string path;
        int dataCounts = 0;
        int taperRatio = 0;
        bool fft_ctrl = false;
        SolidBrush semiTransBrush = new SolidBrush(Color.FromArgb(100, Color.YellowGreen));//70
        SolidBrush semiTransBrush2 = new SolidBrush(Color.FromArgb(200, Color.Black));//200

        #endregion

        public MainForm()
        {            
            InitializeComponent();
        }     

        private void Real_Time_Stop_Button(object sender, EventArgs e)
        {
            if(sPort!=null)
                if(sPort.IsOpen)
                {
                    sPort.Write("false");
                    isCont = false;
                    System.Threading.Thread.Sleep(1000);
                    if (mainThread.IsAlive)
                        mainThread.Abort();

                    monitoringGraphics.Dispose();

                    #region clearVariables
                    data = null;
                    dataArray = null;
                    signal1 = null;
                    signal2 = null;
                    signal3 = null;
                    #endregion

                    #region Enable Disable Controls
                    dataRecording_gbox.Enabled = true;
                    File_gbox.Enabled = true;
                    Live_Button.Enabled = true;
                    Stop_Button1.Enabled = false;
                    Disconnect_Button.Enabled = true;
                    File_Button.Enabled = true;
                    #endregion
                }            
        }
        
        private void Record_Start_Button(object sender, EventArgs e)
        {
            

            #region initializeComponents
            Real_Time_Pbox.BackColor = Color.Black;
            Filtered_Pbox.BackColor = Color.Black;
            Filtered_Pbox.Dock = DockStyle.Fill;
            timePanel.BackColor = Color.Red;
            timeLabel.BackColor = Color.Red;
            timeLabel.Text = delayTime_nup.Value.ToString();
            timeLabel.BackColor = Color.Red;
            #endregion

            #region Enable Disable Controls
            dataMonitoring_gbox.Enabled = false;
            File_gbox.Enabled = false;
            Record_Start_Buton.Enabled = false;
            Record_Stop_Button.Enabled = true;
            time_Nud.Enabled = false;
            delayTime_nup.Enabled = false;
            Disconnect_Button.Enabled = false;
            zoom_Inc.Enabled = false;
            zoom_Dec.Enabled = false;
            fullScreen.Enabled = false;
            File_Button.Enabled = false;
            Save_Button.Enabled = false;
            analysisButton.Enabled = false;
            Plot_Button.Enabled = false;
            startAnalysis_Button.Enabled = false;
            V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            HV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            real_FFT_V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            real_FFT_NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            real_FFT_EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            #endregion

            #region clearComponents_and_Variables
            V_Chart.Series[0].Points.Clear();
            NS_Chart.Series[0].Points.Clear();
            EW_Chart.Series[0].Points.Clear();
            smoothedV_Chart.Series[0].Points.Clear();
            smoothedNS_Chart.Series[0].Points.Clear();
            smoothedEW_Chart.Series[0].Points.Clear();
            HV_Chart.Series.Clear();
            real_FFT_V_Chart.Series.Clear();
            real_FFT_NS_Chart.Series.Clear();
            real_FFT_EW_Chart.Series.Clear();

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            dataGridView1.Rows.Clear();

            filename = null;
            FileName.Text = "...";
            Real_Time_Pbox.Image = null;
            Filtered_Pbox.Image = null;

            #endregion

            #region initializeVariables
            stopTime = time_Nud.Value * 60 * 1000;
            int backCount = (int)delayTime_nup.Value;
            isCont = true;
            int cnt = 0;
            bool start = false;
            #endregion

            #region clear_andInitializeSerialPort
            sPort.DiscardInBuffer();
            sPort.DiscardOutBuffer();  
            sPort.DiscardInBuffer();
            sPort.DiscardOutBuffer();
            sPort.DiscardInBuffer();
            sPort.DiscardOutBuffer();
            data = sPort.ReadLine();
            sPort.Write("true");
            #endregion

            sw = new StreamWriter("datam.txt");  
            
            while (isCont)
            {
                data = sPort.ReadLine().ToString();
                cnt++;
                if (cnt % 200 == 0)
                {
                    backCount--;
                    if (backCount == 0)
                        timeLabel.Text = (time_Nud.Value * 60).ToString();
                    else
                        timeLabel.Text = backCount.ToString();

                    if (backCount == 0 && !start)
                    {
                        start = true;
                        backCount = (int)time_Nud.Value * 60;
                        timeLabel.BackColor = Color.GreenYellow;
                        timePanel.BackColor = Color.GreenYellow;
                    }
                }

                if (cnt >= (delayTime_nup.Value * 200))
                {
                    string[] d = data.Split('*');
                    if(d.Length>2)
                        sw.WriteLine(d[0] + "\t" + d[1] + "\t" + d[2]);
                }

                if (cnt >= (time_Nud.Value * 60 * 1000 / 5 + delayTime_nup.Value * 200))
                {
                    isCont = false;
                    sw.Close();
                    Record_Start_Buton.Enabled = true;
                    Record_Stop_Button.Enabled = false;
                    timePanel.BackColor = Color.Red;
                    timeLabel.BackColor = Color.Red;
                    Save_Button.Enabled = true;
                    Plot_Button.Enabled = true;
                    delayTime_nup.Enabled = true;
                    #region Enable Disable Controls
                    dataMonitoring_gbox.Enabled = true;
                    File_gbox.Enabled = true;
                    Record_Start_Buton.Enabled = true;
                    Record_Stop_Button.Enabled = false;
                    time_Nud.Enabled = true;
                    Disconnect_Button.Enabled = true;
                    File_Button.Enabled = true;
                    #endregion

                    sPort.Write("false");
                    
                }
                Application.DoEvents();
            }
            for (int bip = 0; bip < 3; bip++)
            {
                Console.Beep();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void Record_Stop_Button2(object sender, EventArgs e)
        {
            isCont = false;
            sw.Close();
            sPort.Write("false");

            timePanel.BackColor = Color.Red;
            timeLabel.BackColor = Color.Red;

            #region EnableDisableControls
            Save_Button.Enabled = true;
            Plot_Button.Enabled = true;
            dataMonitoring_gbox.Enabled = true;
            File_gbox.Enabled = true;            
            Record_Start_Buton.Enabled = true;
            Record_Stop_Button.Enabled = false;
            time_Nud.Enabled = true;
            delayTime_nup.Enabled = true;
            Disconnect_Button.Enabled = true;
            File_Button.Enabled = true;
            #endregion
        }

        private void loadComPort()
        {
            Port_Combo.Items.Clear();
            string[] ports = System.IO.Ports.SerialPort.GetPortNames();
            foreach (string port in ports)
                Port_Combo.Items.Add(port);
            Port_Combo.Items.Insert(0, "SELECT");
            Port_Combo.SelectedIndex = 0;
        }

        private void Plot_Button_Click(object sender, EventArgs e)
        {           

            try
            {
                #region initializeComponents
                Real_Time_Pbox.Dock = DockStyle.Left;
                Real_Time_Pbox.BackColor = Color.Black;
                Filtered_Pbox.BackColor = Color.Black;
                panel3.BackColor = Color.Black;

                windowLen_cb.SelectedIndex = 0;
                Real_Time_Pbox.Image = null;
                Filtered_Pbox.Image = null;

                V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

                V_Chart.Series[0].Points.Clear();
                NS_Chart.Series[0].Points.Clear();
                EW_Chart.Series[0].Points.Clear();
                lat_lbl.Text = "...";
                lon_lbl.Text = "...";
                process_pbox.Visible = true;

                filename = "datam.txt";

                double avgNS = 0, avgEW = 0, avgV = 0;
                double tim = 0;
                cn = 0.1;
                #endregion

                font = new Font("Arial", 10, FontStyle.Bold);

                fileNameWithoutExtension = "datam";
                directoryName = Application.StartupPath;
                fullName = directoryName + "\\" + fileNameWithoutExtension + ".info";

                StreamReader sr = new StreamReader(filename);

                dataCount = File.ReadAllLines(filename).Count();
                orginalSignal = new double[4, dataCount];
                orginalSignalRev = new double[4, dataCount];
                orginalSignalRevCopy = new double[4, dataCount];//Trend effect removed data

                for (int i = 0; i < dataCount; i++)
                {
                    Application.DoEvents();
                    string all = sr.ReadLine();
                    string[] bil = all.Split('\t');

                    orginalSignal[0, i] = tim;
                    orginalSignal[1, i] = double.Parse(bil[0]);
                    orginalSignal[2, i] = double.Parse(bil[1]);
                    orginalSignal[3, i] = double.Parse(bil[2]);

                    avgNS += orginalSignal[1, i];
                    avgEW += orginalSignal[2, i];
                    avgV += orginalSignal[3, i];

                    tim += 0.005;
                }

                sr.Close();

                avgNS /= orginalSignal.GetLength(1);
                avgEW /= orginalSignal.GetLength(1);
                avgV /= orginalSignal.GetLength(1);

                //Remove Trend Effect
                for (int i = 0; i < dataCount; i++)
                {
                    orginalSignalRev[1, i] = orginalSignal[1, i] - avgNS;
                    orginalSignalRev[2, i] = orginalSignal[2, i] - avgEW;
                    orginalSignalRev[3, i] = orginalSignal[3, i] - avgV;
                }

                orginalSignalRevCopy = orginalSignalRev;               
                
                drawSignal(orginalSignalRev, 0);

                #region initializeComponents

                if (filteredSignal_Bitmap != null)
                {
                    filteredSignal_Bitmap = null;
                    filteredSignal_Graphics.Clear(Color.Black);
                    filteredSignal_Graphics = null;
                }

                analysisButton.Enabled = true;
                startAnalysis_Button.Enabled = false;
                zoom_Inc.Enabled = true;
                zoom_Dec.Enabled = true;
                process_pbox.Visible = false;
                fullScreen.Enabled = true;
                lowpass_cb.Checked = false;
                filter = false;
                tabControl1.SelectedIndex = 0;
                listBox1.Items.Clear();
                listBox2.Items.Clear();
                #endregion

            }
            catch (Exception ex)
            {                
                process_pbox.Visible = false;
                MessageBox.Show("An error occured during the process, Please, check the data file format");
            }

        }      

        private void File_Button_Click(object sender, EventArgs e)
        {
            try
            {
                #region  initializeComponents
                Real_Time_Pbox.BackColor = Color.Black;
                Filtered_Pbox.BackColor = Color.Black;
                panel3.BackColor = Color.Black;
                windowLen_cb.SelectedIndex = 0;
                Real_Time_Pbox.Image = null;
                Filtered_Pbox.Image = null;
                #endregion

                V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                HV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                real_FFT_V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                real_FFT_NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                real_FFT_EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

                V_Chart.Series[0].Points.Clear();
                NS_Chart.Series[0].Points.Clear();
                EW_Chart.Series[0].Points.Clear();
                smoothedV_Chart.Series[0].Points.Clear();
                smoothedNS_Chart.Series[0].Points.Clear();
                smoothedEW_Chart.Series[0].Points.Clear();
                //real_FFT_V_Chart.Series[0].Points.Clear();
                //real_FFT_NS_Chart.Series[0].Points.Clear();
                //real_FFT_EW_Chart.Series[0].Points.Clear();

                HV_Chart.Series.Clear();

                HV_Chart.Titles[0].Text = null;

                lat_lbl.Text = "...";
                lon_lbl.Text = "...";
                process_pbox.Visible = true;

                listBox1.Items.Clear();
                listBox2.Items.Clear();
                dataGridView1.Rows.Clear();                

                font = new Font("Arial", 10, FontStyle.Bold);                
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "txt files (*.txt)|*.txt";

                //Opening the dialogbox and read data
                if (op.ShowDialog() == DialogResult.OK)
                {                   
                    filename = op.FileName;
                    StreamReader sr = new StreamReader(filename);
                    FileName.Text = filename;

                    dataCount = File.ReadAllLines(filename).Count();
                    orginalSignal = new double[4, dataCount]; 
                    orginalSignalRev = new double[4, dataCount];
                    orginalSignalRevCopy = new double[4, dataCount];

                    double avgNS = 0, avgEW = 0, avgV = 0;
                    double tim = 0;
                    for (int i = 0; i < dataCount; i++)
                    {
                        string all = sr.ReadLine();
                        string[] bil = all.Split('\t');

                        orginalSignal[0, i] = tim;
                        orginalSignal[1, i] = double.Parse(bil[0]);
                        orginalSignal[2, i] = double.Parse(bil[1]);
                        orginalSignal[3, i] = double.Parse(bil[2]);

                        avgNS += orginalSignal[1, i];
                        avgEW += orginalSignal[2, i];
                        avgV += orginalSignal[3, i];

                        tim += 0.005;                      
                    }
                    sr.Close();

                    avgNS /= orginalSignal.GetLength(1);
                    avgEW /= orginalSignal.GetLength(1);
                    avgV /= orginalSignal.GetLength(1);

                    //Remove trend efect
                    for (int i = 0; i < dataCount; i++)
                    {
                        orginalSignalRev[1, i] = orginalSignal[1, i] - avgNS;
                        orginalSignalRev[2, i] = orginalSignal[2, i] - avgEW;
                        orginalSignalRev[3, i] = orginalSignal[3, i] - avgV;
                    }

                    orginalSignalRevCopy = orginalSignalRev;

                    fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(op.FileName);
                    directoryName = System.IO.Path.GetDirectoryName(op.FileName);
                    fullName = directoryName +"\\"+ fileNameWithoutExtension + ".info";
                    if (File.Exists(fullName))
                    { 
                        sr=new StreamReader(fullName);
                        string[] str = sr.ReadLine().Split('\t');
                        if (str[0] == "Latitude:")
                            lat_lbl.Text = str[1];
                        str = sr.ReadLine().Split('\t');
                        if (str[0] == "Longitude:")
                            lon_lbl.Text = str[1];
                        sr.Close();
                    }

                    cn = (float)Real_Time_Pbox.Width/4096.0;
                    drawSignal(orginalSignalRev,0);

                    if (filteredSignal_Bitmap != null)
                    {
                        filteredSignal_Bitmap = null;
                        filteredSignal_Graphics.Clear(Color.Black);
                        filteredSignal_Graphics = null;                        
                    }

                    #region  initializeComponents
                    analysisButton.Enabled = true;
                    startAnalysis_Button.Enabled = false;
                    zoom_Inc.Enabled = true;
                    zoom_Dec.Enabled = true;
                    process_pbox.Visible = false;
                    fullScreen.Enabled = true;
                    lowpass_cb.Checked = false;
                    filter = false;
                    tabControl1.SelectedIndex = 0;
                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    #endregion
                }
                else
                {
                    process_pbox.Visible = false;
                }
            }
            catch(Exception ex)
            {
                process_pbox.Visible = false;
                MessageBox.Show("An error occured during the process, Please, check the data file format");
            } 
        }

        private void drawSignal(double[,] orginalSignal,int type)
        {
            offsetVal();

            Pen pen1 = new Pen(Color.GreenYellow , 0.1f);
            Pen pen2 = new Pen(Color.Red, 0.1f);
            Pen pen3 = new Pen(Color.Turquoise, 0.1f);
            float[] dashValues = { 4, 2 };
            Pen pen4 = new Pen(Color.White, 1);
            pen4.DashPattern = dashValues;
            int timestring = 30;
            if (dataCount > 0)
            {
                if (type == 0)
                {
                    Real_Time_Pbox.Dock = DockStyle.Left;
                    Real_Time_Pbox.Width = (int)(cn * dataCount);
                    orginalSignal_Bitmap = new Bitmap(Real_Time_Pbox.Width, Real_Time_Pbox.Height);
                    Real_Time_Pbox.Image = orginalSignal_Bitmap;
                    orginalSignal_Graphics = Graphics.FromImage(orginalSignal_Bitmap);

                    int timeInt = 0;
                    int timeTrace = 0;

                    if (dataCount < 60000)//10 min
                    {
                        timeTrace = 6000; timeInt = 30;
                    }
                    else if (dataCount < 180000)//30 min
                    {
                        timeTrace = 18000; timeInt = 90;
                    }
                    else if (dataCount < 360000)//60 min
                    {
                        timeTrace = 36000; timeInt = 180;
                    }
                    else
                    {
                        timeTrace = 60000; timeInt = 300;
                    }

                    timestring = timeInt;

                    float x = 0;
                    point1 = orginalSignal[1, 0] / 20.0 + offset1;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = orginalSignal[1, i] / 20.0 + offset1;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        orginalSignal_Graphics.DrawLine(pen1, p1, p2);
                        point1 = point2;

                        if (i % timeTrace == 0)
                        {
                            PointF pf = new PointF(p1.X + 5, 10);

                            orginalSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                            if (windowLen_cb.SelectedIndex>0)
                                orginalSignal_Graphics.DrawString(timestring.ToString() + " sec", new System.Drawing.Font(FontFamily.GenericSansSerif, 10), Brushes.Black, pf);
                            else
                                orginalSignal_Graphics.DrawString(timestring.ToString() + " sec", new System.Drawing.Font(FontFamily.GenericSansSerif, 10), Brushes.White, pf);

                            timestring += timeInt;
                            orginalSignal_Graphics.DrawLine(pen4, p1.X, 0, p1.X, Real_Time_Pbox.Height);
                        }
                    }

                    orginalSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    if (windowLen_cb.SelectedIndex > 0)
                        orginalSignal_Graphics.DrawString("Vertical", font, Brushes.Black, 10, 10);
                    else
                        orginalSignal_Graphics.DrawString("Vertical", font, Brushes.White, 10, 10);

                    x = 0;
                    point1 = orginalSignal[2, 0] / 20.0 + offset2;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = orginalSignal[2, i] / 20.0 + offset2;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        orginalSignal_Graphics.DrawLine(pen2, p1, p2);
                        point1 = point2;
                    }

                    orginalSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    if (windowLen_cb.SelectedIndex > 0)
                        orginalSignal_Graphics.DrawString("North-South", font, Brushes.Black, 10, 198);
                    else
                        orginalSignal_Graphics.DrawString("North-South", font, Brushes.White, 10, 198);

                    x = 0;
                    point1 = orginalSignal[3, 0]/20.0+ offset3;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = orginalSignal[3, i]/20.0+ offset3;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        orginalSignal_Graphics.DrawLine(pen3, p1, p2);
                        point1 = point2;
                    }

                    orginalSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    if (windowLen_cb.SelectedIndex > 0)
                        orginalSignal_Graphics.DrawString("East-West", font, Brushes.Black, 10, 386);
                    else
                        orginalSignal_Graphics.DrawString("East-West", font, Brushes.White, 10, 386);

                    orginalSignal_Bitmap = new Bitmap(Real_Time_Pbox.Width, Real_Time_Pbox.Height, orginalSignal_Graphics);

                    Real_Time_Pbox.Refresh();

                    orginalSignal_Bitmap.Dispose();
                }
                else if (type == 1)
                {
                    Filtered_Pbox.Dock = DockStyle.Left;
                    Filtered_Pbox.Width = (int)(cn * dataCount);
                    filteredSignal_Bitmap = new Bitmap(Filtered_Pbox.Width, Filtered_Pbox.Height);
                    Filtered_Pbox.Image = filteredSignal_Bitmap;
                    filteredSignal_Graphics = Graphics.FromImage(filteredSignal_Bitmap);

                    int timeInt = 0;
                    int timeTrace = 0;

                    if (dataCount < 60000)//10 min
                    {
                        timeTrace = 6000; timeInt = 30;
                    }
                    else if (dataCount < 180000)//30 min
                    {
                        timeTrace = 18000; timeInt = 90;
                    }
                    else if (dataCount < 360000)//60 min
                    {
                        timeTrace = 36000; timeInt = 180;
                    }
                    else
                    {
                        timeTrace = 42000; timeInt = 1200;
                    }

                    timestring = timeInt;

                    float x = 0;
                    point1 = filteredSignal[1, 0] / 10.0+94;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = filteredSignal[1, i] / 10.0+94;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        filteredSignal_Graphics.DrawLine(pen1, p1, p2);
                        point1 = point2;

                        if (i % timeTrace == 0)
                        {
                            PointF pf = new PointF(p1.X + 5, 10);

                            filteredSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                            if (windowLen_cb.SelectedIndex > 0)
                                filteredSignal_Graphics.DrawString(timestring.ToString() + " sec", new System.Drawing.Font(FontFamily.GenericSansSerif, 10), Brushes.Black, pf);
                            else
                                filteredSignal_Graphics.DrawString(timestring.ToString() + " sec", new System.Drawing.Font(FontFamily.GenericSansSerif, 10), Brushes.White, pf);

                            timestring += timeInt;

                            filteredSignal_Graphics.DrawLine(pen4, p1.X, 0, p1.X, Real_Time_Pbox.Height);
                        }

                    }

                    x = 0;
                    point1 = filteredSignal[2, 0] / 10.0+282;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = filteredSignal[2, i] / 10.0+282;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        filteredSignal_Graphics.DrawLine(pen2, p1, p2);
                        point1 = point2;
                    }

                    x = 0;
                    point1 = filteredSignal[3, 0] / 10.0+470;
                    for (int i = 1; i < dataCount; i++)
                    {
                        point2 = filteredSignal[3, i] / 10.0+470;
                        p1 = new PointF(x, (float)point1);
                        x += (float)cn;
                        p2 = new PointF(x, (float)point2);
                        filteredSignal_Graphics.DrawLine(pen3, p1, p2);
                        point1 = point2;
                    }

                    filteredSignal_Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    filteredSignal_Graphics.DrawString("Vertical", font, Brushes.White, 10, 10);
                    filteredSignal_Graphics.DrawString("North-South", font, Brushes.White, 10, (float)(offset2 - offset1 + 10));
                    filteredSignal_Graphics.DrawString("East-West", font, Brushes.White, 10, (float)(offset3 - offset1 + 10));

                    filteredSignal_Bitmap = new Bitmap(Filtered_Pbox.Width, Filtered_Pbox.Height, filteredSignal_Graphics);
                    Filtered_Pbox.Refresh();
                    filteredSignal_Bitmap.Dispose();
                }
            }
        }

        private void window_set()
        {
            cn2 = (float)cn * (float)Convert.ToDouble(windowLen_cb.Text) * 200;
            int array = dataCount / dataCounts;
             
            window_array = new bool[array];
            koord_window = new float[2, array];
            koord_window2 = new float[2, array];

            int i = 0;
            while (i < window_array.Length)
            {
                window_array[i] = true;
                i++;
            }
        }

        private void analysisButton_Click(object sender, EventArgs e)
        {
            windowLen_cb.SelectedIndex = 0;
            tabControl1.SelectedIndex = 0;
            AnalysisPanel.Location = new Point(371,153);
            AnalysisPanel.Visible = true;
            AnalysisPanel.BringToFront();
        }

        private void Save_Button_Click(object sender, EventArgs e)
        {
            coordPanel.Location = new Point(237,161);
            coordPanel.Visible = true;
            tabControl1.SelectedIndex = 0;
        }

        private void zoom_Inc_Click(object sender, EventArgs e)
        {
            if (Real_Time_Pbox.Width < 65535)
            {
                cn *= 2;
                if(tabControl1.SelectedIndex==0)
                    drawSignal(orginalSignalRev,0);
                else if(tabControl1.SelectedIndex==1)
                    drawSignal(orginalSignalRev, 1);
            }
            else
            {
                zoom_Inc.Enabled = false;
                zoom_Dec.Enabled = true;
            }

            if (Real_Time_Pbox.Width >= 600)            
                zoom_Dec.Enabled = true;            
        }

        private void zoom_Dec_Click(object sender, EventArgs e)
        {
            if (Real_Time_Pbox.Width >= 600)
            {
                cn /= 2;
                if (tabControl1.SelectedIndex == 0)
                    drawSignal(orginalSignalRev, 0);
                else if (tabControl1.SelectedIndex == 1)
                    drawSignal(orginalSignalRev, 1);
            }
            else
            {
                zoom_Inc.Enabled = true;
                zoom_Dec.Enabled = false;
            }

            if (Real_Time_Pbox.Width < 65535)
                zoom_Inc.Enabled = true;
        }

        private void fullScreen_Click(object sender, EventArgs e)
        {
            Real_Time_Pbox.Width = this.Width-30;
            cn = Real_Time_Pbox.Width/(float)dataCount;
            if (tabControl1.SelectedIndex == 0)
                drawSignal(orginalSignalRev, 0);
            else if (tabControl1.SelectedIndex == 1)
                drawSignal(orginalSignalRev, 1);
        }

        private void Real_Time_Pbox_MouseClick(object sender, MouseEventArgs e)
        {
            if(windowLen_cb.SelectedIndex>0)
            {
                Real_Time_Pbox.BackColor = Color.White;
                Filtered_Pbox.BackColor = Color.White;
                panel3.BackColor = Color.White;

                float koor = e.X;
                int ii = 0;
                while (ii < window_array.Length)
                {
                    if (koor >= koord_window[0, ii] && koor <= koord_window[1, ii])
                    {
                        if (window_array[ii])
                        {
                            window_array[ii] = false;
                        }
                        else
                        {
                            window_array[ii] = true;
                        }
                    }
                    ii++;
                }
                drawSignal(orginalSignalRev, 0);
                draw_window(orginalSignalRev, Real_Time_Pbox, orginalSignal_Bitmap, orginalSignal_Graphics);
            }
        }

        private void processCancel_button_Click(object sender, EventArgs e)
        {
            AnalysisPanel.Visible = false;
        }

        private void Filtered_Pbox_MouseClick(object sender, MouseEventArgs e)
        {
            if (windowLen_cb.SelectedIndex > 0)
            {
                Real_Time_Pbox.BackColor = Color.White;
                Filtered_Pbox.BackColor = Color.White;
                panel3.BackColor = Color.White;

                float koor = e.X;
                int ii = 0;
                while (ii < window_array.Length)
                {
                    if (koor >= koord_window[0, ii] && koor <= koord_window[1, ii])
                    {
                        if (window_array[ii])
                        {
                            window_array[ii] = false;
                        }
                        else
                        {
                            window_array[ii] = true;
                        }
                    }
                    ii++;
                }

                drawSignal(filteredSignal, 1);
                draw_window(filteredSignal, Filtered_Pbox, filteredSignal_Bitmap, filteredSignal_Graphics);
            }
        }

        private void analysisVoid(double[,] sig)
        {
            #region initializeComponents
            V_Chart.Visible = false;
            NS_Chart.Visible = false;
            EW_Chart.Visible = false;
            HV_Chart.Visible = false;
            #endregion

            #region initializeVariables
            int n = window_array.Length;
            double[,] new_signalNS = new double[n, dataCounts];
            double[,] new_signalEW = new double[n, dataCounts];
            double[,] new_signalUD = new double[n, dataCounts];
            double[,] tapering_signalNS = new double[n, dataCounts];
            double[,] tapering_signalEW = new double[n, dataCounts];
            double[,] tapering_signalUD = new double[n, dataCounts];
            FFTnew_signalNS = new double[n, dataCounts / 2 + 1];
            FFTnew_signalEW = new double[n, dataCounts / 2 + 1];
            FFTnew_signalUD = new double[n, dataCounts / 2 + 1];
            smoothedFFTnew_signalNS = new double[n, dataCounts / 2 + 1];
            smoothedFFTnew_signalEW = new double[n, dataCounts / 2 + 1];
            smoothedFFTnew_signalUD = new double[n, dataCounts / 2 + 1];
            HV = new double[n, dataCounts / 2 + 1];
            HV_Avg = new double[dataCounts / 2 + 1];
            HV_Std = new double[dataCounts / 2 + 1];
            HV_StdMult = new double[dataCounts / 2 + 1];
            HV_StdDiv = new double[dataCounts / 2 + 1];
            #endregion

            //Create Solution File     
            if (fileNameWithoutExtension != "" && directoryName != "")
                swrite = new StreamWriter(directoryName + "//" + fileNameWithoutExtension + "_Solution.Soln");
            else
                swrite = new StreamWriter("temporary_Solution.txt");
            if(filter)
                swrite.WriteLine("Butterworth cutoff frequency: "+frequency);
            if(taper)
                swrite.WriteLine("Cosine taper : %"+ taperRatio);

            int window_count = 0;
            for (int dim = 0; dim < n; dim++)//window
            {
                if (window_array[dim] == true)
                    window_count++;
            }

            swrite.WriteLine("Number of windows:"+window_count);
            swrite.WriteLine("Window length: "+windowLen_cb.Text);
            swrite.WriteLine("Konno-Ohmachi band width: "+bandwidth);            

            //Create Time Windows
            int y = (int)(koord_window2[0, 0]);
            for (int dim = 0; dim < n; dim++)
            {
                if (window_array[dim])
                {
                    for (int c = 0; c < dataCounts; c++)
                    {
                        new_signalUD[dim, c] = sig[1, y];
                        new_signalNS[dim, c] = sig[2, y];
                        new_signalEW[dim, c] = sig[3, y];
                        y++;
                    }
                }
                y = (int)(koord_window2[1, dim]);                
            }

            #region writeFile

            swrite.WriteLine("Time Windows(V)");
            for (int c = 0; c < dataCounts; c++)
            {
                swrite.Write(c + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if(window_array[dim]==true)
                    swrite.Write(new_signalUD[dim, c]+"\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            swrite.WriteLine("Time Windows(NS)");
            for (int c = 0; c < dataCounts; c++)
            {
                swrite.Write(c + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(new_signalNS[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            swrite.WriteLine("Time Windows(EW)");
            for (int c = 0; c < dataCounts; c++)
            {
                swrite.Write(c + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(new_signalEW[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }
            #endregion

            double Fs = 200;//Sampling Frequency
            double T = 1.0 / Fs;//Period
            double[] signal1 = new double[dataCounts];//V
            double[] signal2 = new double[dataCounts];//NS
            double[] signal3 = new double[dataCounts];//EW
            double[] t = new double[n];

            f1 = new double[dataCounts / 2 + 1];
            for (int ii = 0; ii < dataCounts / 2 + 1; ii++)
            {
                f1[ii] = Fs * ii / dataCounts;
            }

            double[] taperWindow = new double[dataCounts];
            if (taper)
            {
                //Cosine Window
                double p = taperRatio / 100.0;
                double C = 1;
                taperWindow = new double[dataCounts];
                for (int zaman = 1; zaman <= dataCounts; zaman++)
                {
                    if (zaman >= 1 && zaman <= (p * dataCounts / 2))
                        taperWindow[zaman - 1] = (C / 2) * (1 - Math.Cos((2 * Math.PI * zaman) / (p * dataCounts + 1)));
                    else if (zaman > (p * dataCounts / 2) && zaman < (dataCounts + 1 - (p * dataCounts / 2)))
                        taperWindow[zaman - 1] = C;
                    else if (zaman >= (dataCounts + 1 - (p * dataCounts / 2)) && zaman <= dataCounts)
                        taperWindow[zaman - 1] = C / 2 * (1 - Math.Cos((2 * Math.PI * (dataCounts + 1 - zaman)) / (p * dataCounts + 1)));
                }
            }
            //Calc FFT
            for (int dim = 0; dim < n; dim++)
            {
                if (window_array[dim])
                {
                    if (taper)
                    {
                        for (int c = 0; c < dataCounts; c++)
                        {
                            signal1[c] = new_signalUD[dim, c] * taperWindow[c];
                            signal2[c] = new_signalNS[dim, c] * taperWindow[c];
                            signal3[c] = new_signalEW[dim, c] * taperWindow[c];
                        }

                        for (int tx = 0; tx < dataCounts; tx++)
                        {
                            tapering_signalUD[dim, tx] = signal1[tx];
                            tapering_signalNS[dim, tx] = signal2[tx];
                            tapering_signalEW[dim, tx] = signal3[tx];
                        }
                    }
                    else
                    {
                        for (int c = 0; c < dataCounts; c++)
                        {
                            signal1[c] = new_signalUD[dim, c];
                            signal2[c] = new_signalNS[dim, c];
                            signal3[c] = new_signalEW[dim, c];
                        }
                    }
                    #region FFT_V
                    //Calculate Fourier Transform
                    //V
                    double[][] Y = new double[dataCounts][];
                    double[] imag = new double[dataCounts];

                    Y = fft_dsa(signal1, imag, dataCounts, 1);

                    double[] P2 = new double[dataCounts];
                    for (int ii = 0; ii < dataCounts; ii++)
                    {
                        P2[ii] = Math.Abs(Y[0][ii] / (dataCounts));
                    }
                    double[] P1 = new double[dataCounts / 2 + 1];

                    for (int ii = 0; ii <= dataCounts / 2; ii++)
                    {
                        P1[ii] = P2[ii];
                    }
                    for (int ii = 1; ii < dataCounts / 2; ii++)
                    {
                        P1[ii] = 2 * P1[ii];
                    }

                    for (int ii = 0; ii < dataCounts / 2; ii++)
                    {
                        FFTnew_signalUD[dim, ii] = P1[ii];
                    }
                    #endregion

                    #region FFT_NS
                    //NS
                    Y = new double[dataCounts][];
                    imag = new double[dataCounts];

                    Y = fft_dsa(signal2, imag, dataCounts, 1);
                    P2 = new double[dataCounts];
                    for (int ii = 0; ii < dataCounts; ii++)
                    {
                        P2[ii] = Math.Abs(Y[0][ii] / (dataCounts));
                    }
                    P1 = new double[dataCounts / 2 + 1];
                    for (int ii = 0; ii <= dataCounts / 2; ii++)
                    {
                        P1[ii] = P2[ii];
                    }
                    for (int ii = 1; ii < dataCounts / 2; ii++)
                    {
                        P1[ii] = 2 * P1[ii];
                    }

                    for (int ii = 0; ii < dataCounts / 2; ii++)
                    {
                        FFTnew_signalNS[dim, ii] = P1[ii];
                    }
                    #endregion

                    #region FFT_EW
                    //EW
                    Y = new double[dataCounts][];
                    imag = new double[dataCounts];

                    Y = fft_dsa(signal3, imag, dataCounts, 1);
                    P2 = new double[dataCounts];
                    for (int ii = 0; ii < dataCounts; ii++)
                    {
                        P2[ii] = Math.Abs(Y[0][ii] / (dataCounts));
                    }
                    P1 = new double[dataCounts / 2 + 1];
                    for (int ii = 0; ii <= dataCounts / 2; ii++)
                    {
                        P1[ii] = P2[ii];
                    }
                    for (int ii = 1; ii < dataCounts / 2; ii++)
                    {
                        P1[ii] = 2 * P1[ii];
                    }

                    for (int ii = 0; ii < dataCounts / 2; ii++)
                    {
                        FFTnew_signalEW[dim, ii] = P1[ii];
                    }
                    ///////////////////////////////////
                    #endregion                    
                    
                    #region Konno_Ohmachi_V
                    //Apply Konno-Ohmachi Smoothing
                    int dataLen = FFTnew_signalUD.GetLength(1);
                    //UD
                    double[] yZ = new double[dataLen];

                    double[] f_shifted = new double[f1.Length];
                    double[] z = new double[dataLen];
                    double[] w = new double[f1.Length];
                    double yyy = 0;

                    for (int iv = 0; iv < f1.Length; iv++)
                    {
                        f_shifted[iv] = f1[iv] / (1 + 1e-4);
                    }

                    //UD Spectrum
                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        Application.DoEvents();
                        if (ix == 0 || ix == dataLen - 1)
                            goto fin;

                        double fc = f1[ix];
                        double[] ww = new double[dataLen];

                        for (int hh = 0; hh < z.Length; hh++)
                        {
                            z[hh] = f_shifted[hh] / fc;
                            ww[hh] = Math.Pow((Math.Sin((double)bandWidth_nup.Value * Math.Log10(z[hh]))) / ((double)bandWidth_nup.Value * Math.Log10(z[hh])), 4);
                            if (double.IsNaN(ww[hh])) ww[hh] = 0;

                            yyy += ww[hh] * FFTnew_signalUD[dim, hh];
                        }

                        yyy = yyy / ww.Sum();

                        yZ[ix] = yyy;
                        yZ[0] = yZ[1];
                        yZ[dataLen - 1] = yZ[dataLen - 2];

                    fin:;
                    }

                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        smoothedFFTnew_signalUD[dim, ix] = yZ[ix];
                    }
                    #endregion

                    #region Konno_Ohmachi_NS
                    //NS
                    double[] yNS = new double[dataLen];

                    f_shifted = new double[f1.Length];
                    z = new double[dataLen];
                    w = new double[f1.Length];
                    yyy = 0;

                    for (int iv = 0; iv < f1.Length; iv++)
                    {
                        f_shifted[iv] = f1[iv] / (1 + 1e-4);
                    }

                    //NS Spectrum
                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        Application.DoEvents();
                        if (ix == 0 || ix == dataLen - 1)
                            goto fin;

                        double fc = f1[ix];
                        double[] ww = new double[dataLen];

                        for (int hh = 0; hh < z.Length; hh++)
                        {
                            z[hh] = f_shifted[hh] / fc;
                            ww[hh] = Math.Pow((Math.Sin((double)bandWidth_nup.Value * Math.Log10(z[hh]))) / ((double)bandWidth_nup.Value * Math.Log10(z[hh])), 4);
                            if (double.IsNaN(ww[hh])) ww[hh] = 0;

                            yyy += ww[hh] * FFTnew_signalNS[dim, hh];
                        }

                        yyy = yyy / ww.Sum();

                        yNS[ix] = yyy;
                        yNS[0] = yNS[1];
                        yNS[dataLen - 1] = yNS[dataLen - 2];

                    fin:;
                    }

                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        smoothedFFTnew_signalNS[dim, ix] = yNS[ix];
                    }
                    #endregion

                    #region Konno_Ohmachi_EW
                    //EW
                    double[] yEW = new double[dataLen];

                    f_shifted = new double[f1.Length];
                    z = new double[dataLen];
                    w = new double[f1.Length];
                    yyy = 0;

                    for (int iv = 0; iv < f1.Length; iv++)
                    {
                        f_shifted[iv] = f1[iv] / (1 + 1e-4);
                    }

                    //EW Spectrum
                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        Application.DoEvents();
                        if (ix == 0 || ix == dataLen - 1)
                            goto fin;

                        double fc = f1[ix];
                        double[] ww = new double[dataLen];

                        for (int hh = 0; hh < z.Length; hh++)
                        {
                            z[hh] = f_shifted[hh] / fc;
                            ww[hh] = Math.Pow((Math.Sin((double)bandWidth_nup.Value * Math.Log10(z[hh]))) / ((double)bandWidth_nup.Value * Math.Log10(z[hh])), 4);
                            if (double.IsNaN(ww[hh])) ww[hh] = 0;

                            yyy += ww[hh] * FFTnew_signalEW[dim, hh];
                        }

                        yyy = yyy / ww.Sum();

                        yEW[ix] = yyy;
                        yEW[0] = yEW[1];
                        yEW[dataLen - 1] = yEW[dataLen - 2];

                    fin:;
                    }

                    for (int ix = 0; ix < dataLen; ix++)
                    {
                        smoothedFFTnew_signalEW[dim, ix] = yEW[ix];
                    }
                    //////////////////////////////////////////////    
                    #endregion

                }
            }
            if (taper)
            {
                #region write_TaperingSignal_FFT_smoothedFFT
                //
                swrite.WriteLine("Tapering Signal V");
                for (int c = 0; c < dataCounts; c++)
                {
                    swrite.Write(c + "\t");
                    for (int dim = 0; dim < n; dim++)
                    {
                        if (window_array[dim] == true)
                            swrite.Write(tapering_signalUD[dim, c] + "\t");
                        else
                            swrite.Write("0" + "\t");
                    }
                    swrite.WriteLine();
                }
                swrite.WriteLine("Tapering Signal NS");
                for (int c = 0; c < dataCounts; c++)
                {
                    swrite.Write(c + "\t");
                    for (int dim = 0; dim < n; dim++)
                    {
                        if (window_array[dim] == true)
                            swrite.Write(tapering_signalNS[dim, c] + "\t");
                        else
                            swrite.Write("0" + "\t");
                    }
                    swrite.WriteLine();
                }
                swrite.WriteLine("Tapering Signal EW");
                for (int c = 0; c < dataCounts; c++)
                {
                    swrite.Write(c + "\t");
                    for (int dim = 0; dim < n; dim++)
                    {
                        if (window_array[dim] == true)
                            swrite.Write(tapering_signalEW[dim, c] + "\t");
                        else
                            swrite.Write("0" + "\t");
                    }
                    swrite.WriteLine();
                }
            }
            #endregion

            //FFT_V
            swrite.WriteLine("FFT V");
            for (int c = 0; c < dataCounts/2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(FFTnew_signalUD[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }
            
            //FFT_NS
            swrite.WriteLine("FFT NS");
            for (int c = 0; c < dataCounts/2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(FFTnew_signalNS[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            //FFT_EW
            swrite.WriteLine("FFT UD");
            for (int c = 0; c < dataCounts/2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(FFTnew_signalEW[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            //Smoothed FFT V
            swrite.WriteLine("smoothed FFT V");
            for (int c = 0; c < dataCounts / 2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(smoothedFFTnew_signalUD[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            //Smoothed FFT NS
            swrite.WriteLine("smoothed FFT NS");
            for (int c = 0; c < dataCounts / 2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(smoothedFFTnew_signalNS[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            //Smoothed FFT EW
            swrite.WriteLine("smoothed FFT EW");
            for (int c = 0; c < dataCounts / 2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(smoothedFFTnew_signalEW[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }            

            #region TimeWindowsListBoxClear
            listBox1.Items.Clear();
            for(int i=0;i<window_array.Length;i++)
            {
                if (window_array[i])
                {
                    listBox1.Items.Add("No:" + (string)(i + 1).ToString("0#"));
                }
                else
                {
                    listBox1.Items.Add("Disable");
                }
            }

            listBox2.Items.Clear();
            for (int i = 0; i < window_array.Length; i++)
            {
                if (window_array[i])
                {
                    listBox2.Items.Add("No:" + (string)(i + 1).ToString("0#"));
                }
                else
                {
                    listBox2.Items.Add("Disable");
                }
            }
            #endregion

            //Calc HV
            for (int dim = 0; dim < n; dim++)
            {
                if (window_array[dim])
                {
                    for (int i = 0; i < smoothedFFTnew_signalNS.GetLength(1); i++)
                    {
                        //HV[dim, i] = Math.Sqrt((Math.Pow(smoothedFFTnew_signalNS[dim, i], 2) + Math.Pow(smoothedFFTnew_signalEW[dim, i], 2))) / smoothedFFTnew_signalUD[dim, i];
                        HV[dim, i] = Math.Sqrt((Math.Pow(smoothedFFTnew_signalNS[dim, i], 2) + Math.Pow(smoothedFFTnew_signalEW[dim, i], 2))/2) / smoothedFFTnew_signalUD[dim, i];
                    }
                }
            }

            swrite.WriteLine("HV Curves");
            for (int c = 0; c < dataCounts / 2; c++)
            {
                swrite.Write(f1[c] + "\t");
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim] == true)
                        swrite.Write(HV[dim, c] + "\t");
                    else
                        swrite.Write("0" + "\t");
                }
                swrite.WriteLine();
            }

            HV_Chart.Series.Clear();
            while (HV_Chart.Series.Count > 0)
            {
                HV_Chart.Series.RemoveAt(0);
            }

            for (int dim = 0; dim < n; dim++)
            {
                if (window_array[dim])
                {
                    Series newSeries = new Series();
                    newSeries.Name = "No:" + dim;
                    newSeries.Palette = ChartColorPalette.Excel;
                    newSeries.BorderWidth = 1;
                    newSeries.ChartType = SeriesChartType.FastLine;

                    HV_Chart.Series.Add(newSeries);
                    for (int i = 0; i < HV.GetLength(1); i++)
                    {
                        if (f1[i] >= 0.1 && f1[i]<=20)
                            newSeries.Points.AddXY(f1[i], HV[dim, i]);
                    }
                }
            }                
            
            int nn = 0;
            for (int i = 0; i < HV.GetLength(1); i++)
            {
                nn = 0;
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim])
                    {
                        HV_Avg[i] += HV[dim, i];
                        nn++;
                    }
                }
                HV_Avg[i] /= nn;
            }

            swrite.WriteLine("Average HV Curve");
            for (int c = 0; c < dataCounts / 2; c++)
            {
                swrite.Write(f1[c] + "\t");
                swrite.Write(HV_Avg[c] + "\t");
                swrite.WriteLine();
            }

            //Std Dev
            for (int i = 0; i < HV.GetLength(1); i++)
            {
                nn = 0;
                for (int dim = 0; dim < n; dim++)
                {
                    if (window_array[dim])
                    {
                        HV_Std[i] += Math.Pow(HV[dim, i] - HV_Avg[i], 2);
                        nn++;
                    }
                }
                HV_Std[i] = Math.Sqrt(HV_Std[i] / (nn - 1));
                HV_StdMult[i] = HV_Avg[i] * Math.Pow(10, HV_Std[i] * Math.Log10(HV_Avg[i]));
                HV_StdDiv[i] = HV_Avg[i] / Math.Pow(10, HV_Std[i] * Math.Log10(HV_Avg[i]));
            }

            /*
            swrite.WriteLine("H/V Standard Deviations");
            for (int i = 0; i < HV.GetLength(1); i++)
            {
                swrite.Write(f1[i] + "\t");
                swrite.Write(HV_Std[i] + "\t");
                swrite.WriteLine();
            }

            swrite.WriteLine("H/V Standard Deviations-Multiplication");
            for (int i = 0; i < HV.GetLength(1); i++)
            {
                swrite.Write(f1[i] + "\t");
                swrite.Write(HV_StdMult[i] + "\t");
                swrite.WriteLine();
            }

            swrite.WriteLine("H/V Standard Deviations-Division");
            for (int i = 0; i < HV.GetLength(1); i++)
            {
                swrite.Write(f1[i] + "\t");
                swrite.Write(HV_StdDiv[i] + "\t");
                swrite.WriteLine();
            }
            */

            swrite.Close();
            
            
            Series newSeries2 = new Series();
            newSeries2.Name = "H/V";
            newSeries2.Palette = ChartColorPalette.None;
            newSeries2.Color = Color.Black;
            newSeries2.BorderWidth = 3;
            newSeries2.ChartType = SeriesChartType.FastLine;

            HV_Chart.Series.Add(newSeries2);
            for (int i = 0; i < HV_Avg.Length; i++)
            {
                if (f1[i] >= 0.1 && f1[i] <= 20)
                    newSeries2.Points.AddXY(f1[i], HV_Avg[i]);
            }
            
            /*
            Series newSeries3 = new Series();
            newSeries3.Name = "H/V*Std";
            newSeries3.Palette = ChartColorPalette.None;
            newSeries3.Color = Color.Red;
            newSeries3.BorderWidth = 1;
            newSeries3.BorderDashStyle = ChartDashStyle.DashDotDot;
            newSeries3.ChartType = SeriesChartType.FastLine;            

            HV_Chart.Series.Add(newSeries3);
            for (int i = 0; i < HV_StdMult.Length; i++)
            {
                if (f1[i] >= 0.1)
                    newSeries3.Points.AddXY(f1[i], HV_StdMult[i]);
            }
            
            Series newSeries4 = new Series();
            newSeries4.Name = "H/V/Std";
            newSeries4.Palette = ChartColorPalette.None;
            newSeries4.Color = Color.Red;
            newSeries4.BorderWidth = 1;
            newSeries4.BorderDashStyle = ChartDashStyle.DashDotDot;
            newSeries4.ChartType = SeriesChartType.FastLine;

            HV_Chart.Series.Add(newSeries4);
            for (int i = 0; i < HV_StdDiv.Length; i++)
            {
                if (f1[i] >= 0.1)
                    newSeries4.Points.AddXY(f1[i], HV_StdDiv[i]);
            }
            */

            HV_Chart.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
            HV_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
            HV_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
            HV_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

            dataGridView1.Rows.Clear();
            int j = 0;
            for (int i = 0; i < HV_Avg.Length; i++)
            {
                if (f1[i] >= 0.1 && f1[i] <= 20)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[j].Cells[0].Value = f1[i];
                    dataGridView1.Rows[j].Cells[1].Value = Math.Round(HV_Avg[i], 6);
                    j++;
                }
            }

            HV_Chart.Annotations.Clear();
            HV_Chart.Titles[0].Text = "";          
        }

        private double[] Konno_Ohmachi_Smoothing(double[] FF1P1,double[] f1)
        {
            double[] yXX = new double[FF1P1.Length];
            double[] f_shifted = new double[f1.Length];
            double[] z = new double[FF1P1.Length];
            double[] w = new double[f1.Length];
            double yyy = 0;

            for (int iv = 0; iv < f1.Length; iv++)
            {
                f_shifted[iv] = f1[iv] / (1 + 1e-4);
            }

            for (int ix = 0; ix < FF1P1.Length; ix++)
            {
                Application.DoEvents();
                if (ix == 0 || ix == FF1P1.Length - 1)
                    goto fin;

                double fc = f1[ix];
                double[] ww = new double[FF1P1.Length];

                for (int hh = 0; hh < z.Length; hh++)
                {
                    z[hh] = f_shifted[hh] / fc;
                    ww[hh] = Math.Pow((Math.Sin((double)bandWidth_nup.Value * Math.Log10(z[hh]))) / ((double)bandWidth_nup.Value * Math.Log10(z[hh])), 4);
                    if (double.IsNaN(ww[hh])) ww[hh] = 0;

                    yyy += ww[hh] * FF1P1[hh];
                }

                yyy = yyy / ww.Sum();

                yXX[ix] = yyy;
                yXX[0] = yXX[1];
                yXX[FF1P1.Length - 1] = yXX[FF1P1.Length - 2];

            fin:;
            }
            return yXX;
        }
       
        private void startAnalysis_Button_Click(object sender, EventArgs e)
        {
            if (windowLen_cb.SelectedIndex > 0)
            {
                #region initializeComponent
                process_pbox.Visible = true;
                processFilter_pb.Visible = true;
                V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                #endregion

                //HV_Chart.ChartAreas[0].AxisY.Minimum = 0;
                //HV_Chart.ChartAreas[0].AxisY.Maximum = 2;

                V_Chart.Series[0].Points.Clear();
                NS_Chart.Series[0].Points.Clear();
                EW_Chart.Series[0].Points.Clear();

                smoothedV_Chart.Series[0].Points.Clear();
                smoothedNS_Chart.Series[0].Points.Clear();
                smoothedEW_Chart.Series[0].Points.Clear();

                //real_FFT_V_Chart.Series[0].Points.Clear();
                //real_FFT_NS_Chart.Series[0].Points.Clear();
                //real_FFT_EW_Chart.Series[0].Points.Clear();                

                if (filter)//Low-pass filtered data
                {
                    Real_Time_Pbox.BackColor = Color.White;
                    Filtered_Pbox.BackColor = Color.White;
                    panel3.BackColor = Color.White;

                    analysisVoid(filteredSignal);

                    V_Chart.Visible = true;
                    NS_Chart.Visible = true;
                    EW_Chart.Visible = true;
                    HV_Chart.Visible = true;
                    tabControl1.SelectedIndex = 4;
                }
                else//Raw Data
                {
                    Real_Time_Pbox.BackColor = Color.White;
                    Filtered_Pbox.BackColor = Color.White;
                    panel3.BackColor = Color.White;

                    analysisVoid(orginalSignalRev);

                    V_Chart.Visible = true;
                    NS_Chart.Visible = true;
                    EW_Chart.Visible = true;
                    HV_Chart.Visible = true;
                    tabControl1.SelectedIndex = 4;
                }
                process_pbox.Visible = false;
                processFilter_pb.Visible = false;
            }
            else
            {
                MessageBox.Show("Please, select the window length");
            }   
        }
      
        private void saveOK_btn_Click(object sender, EventArgs e)
        {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    activeDirectory = folderBrowserDialog1.SelectedPath;
                }

                if (activeDirectory != "")
                {
                    path = activeDirectory + "//" + DateTime.Now.ToString().Replace(" ", "_").Replace(":", "_") ;
                    StreamReader srr = new StreamReader("datam.txt");
                    int count = File.ReadAllLines("datam.txt").Count();
                    if (count > 0)
                    {
                        sw = new StreamWriter(path+ ".txt");
                        while (srr.Peek() != -1)
                        {
                            string bil = srr.ReadLine();
                            string[] bilgi = bil.Split('\t');
                            sw.WriteLine(bilgi[0] + "\t" + bilgi[1] + "\t" + bilgi[2]);
                        }
                        sw.Close();
                        srr.Close();
                        MessageBox.Show("The process has implemented successfully.");
                    }
                    else
                    {
                        MessageBox.Show("The data can not be found");
                    }

                }
                else
                {
                    MessageBox.Show("Please, create or select a project");
                }

                if (latitude_tx.Text != string.Empty && longitude_tx.Text != string.Empty)
                {
                    sw = new StreamWriter(path+".info");
                    sw.WriteLine("Latitude:\t"+latitude_tx.Text);
                    sw.WriteLine("Longitude:\t" + longitude_tx.Text);
                    sw.WriteLine("Description:\t" + description_tx.Text);
                    sw.Close();
                }               
               
                coordPanel.Visible = false;
        }      

        private void saveCANCEL_btn_Click(object sender, EventArgs e)
        {
            coordPanel.Visible = false;
        }

        public static double[][] fft_dsa(double[] fr, double[] fi, int n, int t)
        {
            double tr, ti, wr, wi;
            double[][] result;
            result = new double[2][];
            result[0] = new double[n];
            result[1] = new double[n];
            int nn, l, istep, el;
            int i, j, m, mr;
            if (t < 0)
                for (i = 0; i < n; i++)
                {
                    fr[i] = fr[i] / n;
                    fi[i] = fi[i] / n;
                }
            mr = 0;
            nn = n - 1;
            for (m = 1; m <= nn; m++)
            {
                l = n;
                do { l /= 2; } while ((mr + l) > nn);
                mr = mr % l + l;
                if (mr > m)
                {
                    tr = fr[m];
                    fr[m] = fr[mr];
                    fr[mr] = tr;
                    ti = fi[m];
                    fi[m] = fi[mr];
                    fi[mr] = ti;
                }
            }
            l = 1;
            while (l < n)
            {
                istep = 2 * l;
                el = l;
                for (m = 1; m <= l; m++)
                {
                    wr = (double)Math.Cos((double)(Math.PI * (1 - m) / el));
                    wi = (double)Math.Cos((double)(Math.PI / 2.0 - Math.PI * (1 - m) / el)) * t;
                    for (i = m; i <= n; i += istep)
                    {
                        j = i + l;
                        tr = wr * fr[j - 1] - wi * fi[j - 1];
                        ti = wr * fi[j - 1] + wi * fr[j - 1];
                        fr[j - 1] = fr[i - 1] - tr;
                        fi[j - 1] = fi[i - 1] - ti;
                        fr[i - 1] = fr[i - 1] + tr;
                        fi[i - 1] = fi[i - 1] + ti;
                    }
                }
                l = istep;
            }
            for (i = 0; i < n; i++)
            {
                result[0][i] = fr[i];
                result[1][i] = fi[i];
            }
            return result;
        }

        bool taper=false;
        private void processOK_Button_Click(object sender, EventArgs e)
        {
            if (lowpass_cb.Checked)
            {
                if (Frequency_tx.Text != "")
                {
                    filter = true;
                    frequency = double.Parse(Frequency_tx.Text);
                }
                else
                {
                    MessageBox.Show("Please, enter the filtering parameter");
                    goto fin;
                }
            }
            else
            {
                filter = false;
            }

            if (taper_cb.Checked)
            {
                if (Frequency_tx.Text != "")
                {
                    taper = true;
                    taperRatio = (int)taperRatio_nup.Value;
                }                
            }
            else
            {
                taper = false;
            }

            if (windowLen_cb.SelectedIndex>0)
            {
                    window = windowLen_cb.SelectedIndex;
                    switch(window)
                    {
                        case 1:
                            dataCounts = 2048;//1024x2 (0.005 ms sampling rate)
                            break;
                        case 2:
                            dataCounts = 4096;//2048x2
                            break;
                        case 3:
                            dataCounts = 8192;
                            break;
                        case 4:
                            dataCounts = 16384;
                            break;
                        case 5:
                            dataCounts = 32768;
                            break;
                }                
            }
            else
            {
                MessageBox.Show("Please, select the window length");
                goto fin;                         
            }

                if (bandWidth_nup.Value.ToString() != "")
                {
                    bandwidth = double.Parse(bandWidth_nup.Value.ToString());
                }
                else
                {
                    MessageBox.Show("Please, enter the smoothing parameter");
                    goto fin;
                }

            

            AnalysisPanel.Visible = false;

            if (filter)
            {
                tabControl1.SelectedIndex = 1;
                processFilter_pb.Visible = true;

               


                if (filename == "")
                {
                    filename = "datam.txt";
                }

                if (File.Exists(filename))
                {
                    StreamReader sr = new StreamReader(filename);
                    int count = File.ReadAllLines(filename).Count();

                    filteredSignal = new double[4, count];
                    filteredSignal = Butterworth(orginalSignalRevCopy,0.005,frequency);
                    
                    sr.Close();

                    offsetVal();

                    drawSignal(filteredSignal,1);

                    startAnalysis_Button.Enabled = true;
                    zoom_Inc.Enabled = true;
                    zoom_Dec.Enabled = true;
                    process_pbox.Visible = false;
                    fullScreen.Enabled = true;
                }
                processFilter_pb.Visible = false;
            }

                Real_Time_Pbox.BackColor = Color.White;
                Filtered_Pbox.BackColor = Color.White;
                panel3.BackColor = Color.White;
                startAnalysis_Button.Enabled = true;

            window_set();

            if (window_array.GetLength(0) == 0)
            {
                MessageBox.Show("Please, select a smaller window size");
                goto fin;
            }
                if (filter)
                {
                    drawSignal(filteredSignal, 1);
                    draw_window(filteredSignal, Filtered_Pbox, filteredSignal_Bitmap, filteredSignal_Graphics);
                }
                else
                {
                    drawSignal(orginalSignalRev, 0);
                    draw_window(orginalSignalRev, Real_Time_Pbox, orginalSignal_Bitmap, orginalSignal_Graphics);
                }

        fin:;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBox1.SelectedItem!=null)
            if (listBox1.SelectedItem.ToString()!="Disable")
            {
                //V
                V_Chart.Series[0].Points.Clear();
                V_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                        if (f1[ii] >= 0.1)
                            V_Chart.Series[0].Points.AddXY(f1[ii],FFTnew_signalUD[listBox1.SelectedIndex,ii]);
                }
                V_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                V_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                V_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;
                
                //NS
                NS_Chart.Series[0].Points.Clear();
                NS_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                        if (f1[ii] >= 0.1)
                            NS_Chart.Series[0].Points.AddXY(f1[ii], FFTnew_signalNS[listBox1.SelectedIndex, ii]);
                }
                NS_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                NS_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;
                
                //EW
                EW_Chart.Series[0].Points.Clear();
                EW_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                        if (f1[ii] >= 0.1)
                            EW_Chart.Series[0].Points.AddXY(f1[ii], FFTnew_signalEW[listBox1.SelectedIndex, ii]);
                }
                EW_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                EW_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;
            }
            else
            {
                V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

                V_Chart.Series[0].Points.Clear();
                NS_Chart.Series[0].Points.Clear();
                EW_Chart.Series[0].Points.Clear();
            }

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBox2.SelectedItem!=null)
            if (listBox2.SelectedItem.ToString() != "Disable")
            {
                //V
                smoothedV_Chart.Series[0].Points.Clear();
                smoothedV_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    smoothedV_Chart.Series[0].Points.AddXY(f1[ii], smoothedFFTnew_signalUD[listBox2.SelectedIndex, ii]);
                }
                smoothedV_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                smoothedV_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

                //NS
                smoothedNS_Chart.Series[0].Points.Clear();
                smoothedNS_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    smoothedNS_Chart.Series[0].Points.AddXY(f1[ii], smoothedFFTnew_signalNS[listBox2.SelectedIndex, ii]);
                }
                smoothedNS_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                smoothedNS_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

                //EW
                smoothedEW_Chart.Series[0].Points.Clear();
                smoothedEW_Chart.Series[1].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    smoothedEW_Chart.Series[0].Points.AddXY(f1[ii], smoothedFFTnew_signalEW[listBox2.SelectedIndex, ii]);
                }
                smoothedEW_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
                smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
                smoothedEW_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;
            }
            else
            {
                smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
                smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

                smoothedV_Chart.Series[0].Points.Clear();
                smoothedNS_Chart.Series[0].Points.Clear();
                smoothedEW_Chart.Series[0].Points.Clear();
            }
        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            /*
            Real_Time_Pbox.Dock = DockStyle.Fill;
            RealTimeDrawArea = new Bitmap(Real_Time_Pbox.Size.Width, Real_Time_Pbox.Size.Height);
            Real_Time_Pbox.Image = RealTimeDrawArea;

            monitoringGraphics = Graphics.FromImage(RealTimeDrawArea);
            */
            offsetVal();
        }

        private void time_Nud_ValueChanged(object sender, EventArgs e)
        {
            timeLabel.Text = (time_Nud.Value * 60).ToString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lowpass_cb_CheckedChanged(object sender, EventArgs e)
        {

        }

       

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
                fft_ctrl = true;
            else
                fft_ctrl = false;
        }
               
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sPort != null)
            {
                if (sPort.IsOpen)
                {
                    sPort.Write("false");
                    isCont = false;
                }
                System.Threading.Thread.Sleep(1000);

                if(mainThread!=null)
                if (mainThread.IsAlive)
                    mainThread.Abort();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(dataGridView1.Rows.Count>1)
            {
                double f1 = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[0].Value);
                double val = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
                int say = 0;
                for(int i=0;i<window_array.Length;i++)
                {
                    if (window_array[i])
                        say++;
                }

                HV_Chart.Titles[0].Text = "Number of Windows: " + say + ", fo:" + Math.Round(f1, 2) + " Hz, Ao=" + Math.Round(val, 2);
                HV_Chart.Annotations.Clear();

                ChartArea CA3 = HV_Chart.ChartAreas[0];
                VerticalLineAnnotation VA3 = new VerticalLineAnnotation();
                VA3.AxisX = CA3.AxisX;
                VA3.AllowMoving = false;
                VA3.IsInfinitive = true;
                VA3.ClipToChartArea = CA3.Name;
                VA3.Name = "myLine";
                VA3.LineColor = Color.Red;
                VA3.LineWidth = 3;
                VA3.LineDashStyle = ChartDashStyle.Dash;
                VA3.X = f1;
                HV_Chart.Annotations.Add(VA3);
            }
        }

        public void draw_window(double[,] signal,PictureBox pbox,Bitmap wBitmap,Graphics Gph)
        {
            wBitmap = new Bitmap(pbox.Width, pbox.Height);
            Gph = Graphics.FromImage(pbox.Image);
            Pen penn = new Pen(Color.Black, 1);
            refer = 0;
            int ii = 0;
            int dtCo = 0;
            while (ii<koord_window.GetLength(1))
            {
                koord_window[0, ii] = refer;
                koord_window2[0, ii] = dtCo;
                if (window_array[ii])
                {
                    Gph.DrawRectangle(penn, refer, 0, cn2, pbox.Height - 5);
                    Gph.FillRectangle(semiTransBrush, refer, 0, cn2, pbox.Height - 5);
                }
                else
                {
                    Gph.DrawRectangle(penn, refer, 0, cn2, pbox.Height - 5);
                    Gph.FillRectangle(semiTransBrush2, refer, 0, cn2, pbox.Height - 5);
                }
                refer += cn2;
                koord_window[1, ii] = refer;
                dtCo += dataCounts;
                koord_window2[1, ii] = dtCo;
                ii++;
            }

            windowing_Bitmap = new Bitmap(pbox.Width, pbox.Height, Gph);
            pbox.Refresh();
            pbox.Cursor = Cursors.Hand;
        }

        public static double[,] Butterworth(double[,] indata, double deltaTimeinsec, double CutOff)
        {
            double[] signal = new double[indata.GetLength(1)];
            double[,] signalfin = new double[4, indata.GetLength(1)];

            if (indata == null) return null;
            if (CutOff == 0) return indata;

            for (int i = 0; i < 3; i++)
            {
                for (int klk = 0; klk < indata.GetLength(1); klk++)
                {
                    signal[klk] = indata[i+1, klk];
                }

                double Samplingrate = 1 / deltaTimeinsec;
                long dF2 = indata.GetLength(1) - 1;        // The data range is set with dF2
                double[] Dat2 = new double[dF2 + 4]; // Array with 4 extra points front and back
                double[] data = signal; // Ptr., changes passed data

                // Copy indata to Dat2
                for (long r = 0; r < dF2; r++)
                {
                    Dat2[2 + r] = signal[r];
                }
                Dat2[1] = Dat2[0] = signal[0];
                Dat2[dF2 + 3] = Dat2[dF2 + 2] = signal[dF2];

                const double pi = 3.14159265358979;
                double wc = Math.Tan(CutOff * pi / Samplingrate);
                double k1 = 1.414213562 * wc; // Sqrt(2) * wc
                double k2 = wc * wc;
                double a = k2 / (1 + k1 + k2);
                double b = 2 * a;
                double c = a;
                double k3 = b / k2;
                double d = -2 * a + k3;
                double e = 1 - (2 * a) - k3;

                // RECURSIVE TRIGGERS - ENABLE filter is performed (first, last points constant)
                double[] DatYt = new double[dF2 + 4];
                DatYt[1] = DatYt[0] = signal[0];
                for (long s = 2; s < dF2 + 2; s++)
                {
                    DatYt[s] = a * Dat2[s] + b * Dat2[s - 1] + c * Dat2[s - 2]
                               + d * DatYt[s - 1] + e * DatYt[s - 2];
                }
                DatYt[dF2 + 3] = DatYt[dF2 + 2] = DatYt[dF2 + 1];

                // FORWARD filter
                double[] DatZt = new double[dF2 + 2];
                DatZt[dF2] = DatYt[dF2 + 2];
                DatZt[dF2 + 1] = DatYt[dF2 + 3];
                for (long t = -dF2 + 1; t <= 0; t++)
                {
                    DatZt[-t] = a * DatYt[-t + 2] + b * DatYt[-t + 3] + c * DatYt[-t + 4]
                                + d * DatZt[-t + 1] + e * DatZt[-t + 2];
                }

                // Calculated points copied for return
                for (long p = 0; p < dF2; p++)
                {
                    data[p] = DatZt[p];
                }

                for (int klk = 0; klk < indata.GetLength(1); klk++)
                {
                    signalfin[i+1,klk] = data[klk];
                }
            }

            return signalfin;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            loadComPort();        
        }      

        private void Refresh_Button_Click(object sender, EventArgs e)
        {
            loadComPort();
        }

        private void Connect_Button_Click(object sender, EventArgs e)
        {
            if (Port_Combo.SelectedIndex > 0)
            {
                try
                {
                    sPort = new SerialPort();
                    sPort.PortName = Port_Combo.Text;
                    sPort.BaudRate = 250000;
                    sPort.Open();

                    sPort.DiscardInBuffer();
                    sPort.DiscardOutBuffer();

                    #region Enable_Disable_Controls
                    Connect_Button.Enabled = false;
                    Disconnect_Button.Enabled = true;
                    dataMonitoring_gbox.Enabled = true;
                    dataRecording_gbox.Enabled = true;
                    Port_Combo.Enabled = false;
                    Refresh_Button.Enabled = false;
                    #endregion

                    toolStripStatusLabel1.Text = "COM: Connected";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The connection is not established");
                }
            }
            else
            {
                MessageBox.Show("Please, select a port");
            }
        }

        private void Disconnect_Button_Click(object sender, EventArgs e)
        {
            if (sPort.IsOpen)
            {
                sPort.Close();

                #region Enable_Disable_Controls
                Connect_Button.Enabled = true;
                Disconnect_Button.Enabled = false;
                dataMonitoring_gbox.Enabled = false;
                dataRecording_gbox.Enabled = false;
                Port_Combo.Enabled = true;
                Refresh_Button.Enabled = true;
                #endregion

                toolStripStatusLabel1.Text = "COM: Disconnected";
            }
        }

        private void Live_Button_Click(object sender, EventArgs e)
        {
            #region initializeVariables
            filename = null;
            FileName.Text = "...";

            if (monitoringGraphics != null)
                monitoringGraphics.Dispose();
            if (RealTimeDrawArea != null)
                RealTimeDrawArea.Dispose();
            if (dataArray != null)
                dataArray = null;
            if (orginalSignal != null)
                orginalSignal = null;
            if (orginalSignalRev != null)
                orginalSignalRev = null;
            if (filteredSignal != null)
                filteredSignal = null;
            if (orginalSignal_Graphics != null)
                orginalSignal_Graphics.Dispose();
            if (filteredSignal_Graphics != null)
                filteredSignal_Graphics.Dispose();
            if (orginalSignal_Bitmap != null)
                orginalSignal_Bitmap.Dispose();
            if (filteredSignal_Bitmap != null)
                filteredSignal_Bitmap.Dispose();
            if (windowing_Bitmap != null)
                windowing_Bitmap.Dispose();
            if (window_array != null)
                window_array = null;
            if (koord_window != null)
                koord_window = null;

            orginalSignal = null;
            filteredSignal = null;
            orginalSignalRev = null;
            #endregion

            #region initializeComponent
            Real_Time_Pbox.Image = null;
            Filtered_Pbox.Image = null;

            Real_Time_Pbox.BackColor = Color.Black;
            Filtered_Pbox.BackColor = Color.Black;

            Filtered_Pbox.Dock = DockStyle.Fill;

            V_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

            smoothedV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            smoothedNS_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;
            smoothedEW_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

            HV_Chart.ChartAreas[0].AxisX.IsLogarithmic = false;

            V_Chart.Series[0].Points.Clear();
            NS_Chart.Series[0].Points.Clear();
            EW_Chart.Series[0].Points.Clear();

            HV_Chart.Series.Clear();
            HV_Chart.Titles[0].Text = null;

            real_FFT_V_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
            real_FFT_V_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
            real_FFT_V_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

            real_FFT_NS_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
            real_FFT_NS_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
            real_FFT_NS_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

            real_FFT_EW_Chart.ChartAreas[0].AxisX.Minimum = 0.1;
            real_FFT_EW_Chart.ChartAreas[0].AxisX.IsLogarithmic = true;
            real_FFT_EW_Chart.ChartAreas[0].AxisX.LogarithmBase = 10;

            Real_Time_Pbox.Dock = DockStyle.Fill;
            RealTimeDrawArea = new Bitmap(Real_Time_Pbox.Size.Width, Real_Time_Pbox.Size.Height);
            Real_Time_Pbox.Image = RealTimeDrawArea;

            monitoringGraphics = Graphics.FromImage(RealTimeDrawArea);
            pen1 = new Pen(Brushes.GreenYellow,0.1f);
            pen2 = new Pen(Brushes.Red, 0.1f);
            pen3 = new Pen(Brushes.Turquoise, 0.1f);
            font = new Font("Arial", 10, FontStyle.Bold);
            #endregion

            mainThread = new Thread(new ThreadStart(mainThreadVoid));
            mainThread.IsBackground = true;

            #region Enable Disable Controls
            dataRecording_gbox.Enabled = false;
            File_gbox.Enabled = false;            
            Live_Button.Enabled = false;
            Stop_Button1.Enabled = true;
            Disconnect_Button.Enabled = false;
            zoom_Inc.Enabled = false;
            zoom_Dec.Enabled = false;
            fullScreen.Enabled = false;
            File_Button.Enabled = false;
            Save_Button.Enabled = false;
            analysisButton.Enabled = false;
            Plot_Button.Enabled = false;
            startAnalysis_Button.Enabled = false;
            #endregion

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            dataGridView1.Rows.Clear();

            signal1 = new double[N];
            signal2 = new double[N];
            signal3 = new double[N];

            signalReal1 = new double[4 * N];
            signalReal2 = new double[4 * N];
            signalReal3 = new double[4 * N];

            offsetVal();

            int i = 0;
            foreach(double item in signalReal1)
            {
                signalReal1[i] = offset1;
                signalReal2[i] = offset2;
                signalReal3[i] = offset3;
                i++;
            }

            isCont = true;
            mainThread.Start();
        }

        public void offsetVal()
        {         
            offset1 = Real_Time_Pbox.Size.Height / 6.0;
            offset2 = Real_Time_Pbox.Size.Height / 2.0;
            offset3 = 5.0 * Real_Time_Pbox.Size.Height / 6.0;
            sampling = (float)Real_Time_Pbox.Size.Width / (float)(4*N);           
        }

        private void mainThreadVoid()
        {            
            int i = 0;
            sPort.Write("true");
            while (isCont)
            {                
                data = sPort.ReadLine().Replace("\r", "");
                if (data != "")
                {
                    dataArray = data.Split('*');
                    try
                    {
                        if (dataArray[0] == "") dataArray[0] = "0";
                        if (dataArray[1] == "") dataArray[1] = "0";
                        if (dataArray[2] == "") dataArray[2] = "0";

                        if (i < N)
                        {
                            signal1[i] = double.Parse(dataArray[0]);
                            signal2[i] = double.Parse(dataArray[1]);
                            signal3[i] = double.Parse(dataArray[2]);

                            if (signal1[i] == 2048)
                                signal1[i] = offset1;
                            else
                                signal1[i] = Math.Round((signal1[i] - 2048) / 10.0, 1) + offset1;

                            if (signal2[i] == 2048)
                                signal2[i] = offset2;
                            else
                                signal2[i] = Math.Round((signal2[i] - 2048) / 10.0, 1) + offset2;

                            if (signal3[i] == 2048)
                                signal3[i] = offset3;
                            else
                                signal3[i] = Math.Round((signal3[i] - 2048) / 10.0, 1) + offset3;

                        }
                        else
                        {
                            //label14.Text = dataArray[0];
                            //label15.Text = dataArray[1];
                            //label16.Text = dataArray[2];

                            Array.Copy(signalReal1, N - 1, signalReal1, 0, 3 * N);
                            Array.Copy(signal1, 0, signalReal1, 3 * N - 1, N);
                            Array.Copy(signalReal2, N - 1, signalReal2, 0, 3 * N);
                            Array.Copy(signal2, 0, signalReal2, 3 * N - 1, N);
                            Array.Copy(signalReal3, N - 1, signalReal3, 0, 3 * N);
                            Array.Copy(signal3, 0, signalReal3, 3 * N - 1, N);

                            monitoringGraphics.Clear(Color.Black);
                            if (!fft_ctrl)
                            {
                                this.Invoke((MethodInvoker)delegate { updateSignalGraphics(); });
                                i = -1;
                            }
                            else
                            {
                                this.Invoke((MethodInvoker)delegate { updateSignalGraphics(); });
                                this.Invoke((MethodInvoker)delegate { updatefftGraphics(); });
                                i = -1;
                            }
                        }
                        i++;
                    
                    }
                    catch
                    { }
                }               
            }
        }       

        private void drawSignal(double[] signal,Pen pencil)
        {
            float x = 0;
            point1 = signal[0];
            foreach (double value in signal)
            {
                point2 = value;
                p1 = new PointF(x, (float)point1);
                x += sampling;
                p2 = new PointF(x, (float)point2);
                monitoringGraphics.DrawLine(pencil, p1, p2);
                point1 = point2;
            }
            signal = null;
        }

        private void updateSignalGraphics()
        {
            drawSignal(signalReal1, pen1);
            drawSignal(signalReal2, pen2);
            drawSignal(signalReal3, pen3);

            monitoringGraphics.DrawString("Vertical", font, Brushes.White, 10, 10);
            monitoringGraphics.DrawString("North-South", font, Brushes.White, 10, (float)(offset2-offset1+10));
            monitoringGraphics.DrawString("East-West", font, Brushes.White, 10, (float)(offset3-offset1+10));

            Real_Time_Pbox.Refresh();

        }

        double[] FFTrealtime_signalV=new double[N/2+1];
        double[] FFTrealtime_signalNS = new double[N / 2 + 1];
        double[] FFTrealtime_signalEW = new double[N / 2 + 1];
        private void updatefftGraphics()
        {           
                dataCounts = N;
                double[] ss = new double[dataCounts];
                ss = signal1;
                #region FFT_V
                //Calculating the Fourier Transform of Data
                //V
                double[][] Y = new double[dataCounts][];
                double[] imag = new double[dataCounts];

                Y = fft_dsa(ss, imag, dataCounts, 1);

                double[] P2 = new double[dataCounts];
                for (int ii = 0; ii < dataCounts; ii++)
                {
                    P2[ii] = Math.Abs(Y[0][ii] / dataCounts);
                }
                double[] P1 = new double[dataCounts / 2 + 1];

                for (int ii = 0; ii <= dataCounts / 2; ii++)
                {
                    P1[ii] = P2[ii];
                }
                for (int ii = 1; ii < dataCounts / 2; ii++)
                {
                    P1[ii] = 2 * P1[ii];
                }

                for (int ii = 0; ii < dataCounts / 2; ii++)
                {
                    FFTrealtime_signalV[ii] = P1[ii];
                }
                #endregion

                double Fs = 200;//Sampling Frequency
                double T = 1.0 / Fs;
                f1 = new double[dataCounts / 2 + 1];
                for (int ii = 0; ii < dataCounts / 2 + 1; ii++)
                {
                    f1[ii] = Fs * ii / dataCounts;
                }
                
                if(real_FFT_V_Chart.Series[0].Points.Count>0)
                    real_FFT_V_Chart.Series[0].Points.Clear();

                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    if (f1[ii] >= 0.1)
                        real_FFT_V_Chart.Series[0].Points.AddXY(f1[ii], FFTrealtime_signalV[ii]);
                }

                //***********************************************************
                dataCounts = N;
                ss = new double[dataCounts];
                ss = signal2;
                #region FFT_NS
                //Calculating the Fourier Transform of Data
                //NS
                Y = new double[dataCounts][];
                imag = new double[dataCounts];

                Y = fft_dsa(ss, imag, dataCounts, 1);

                P2 = new double[dataCounts];
                for (int ii = 0; ii < dataCounts; ii++)
                {
                    P2[ii] = Math.Abs(Y[0][ii] / dataCounts);
                }
                P1 = new double[dataCounts / 2 + 1];

                for (int ii = 0; ii <= dataCounts / 2; ii++)
                {
                    P1[ii] = P2[ii];
                }
                for (int ii = 1; ii < dataCounts / 2; ii++)
                {
                    P1[ii] = 2 * P1[ii];
                }

                for (int ii = 0; ii < dataCounts / 2; ii++)
                {
                    FFTrealtime_signalNS[ii] = P1[ii];
                }
                #endregion

                f1 = new double[dataCounts / 2 + 1];
                for (int ii = 0; ii < dataCounts / 2 + 1; ii++)
                {
                    f1[ii] = Fs * ii / dataCounts;
                }
                
                if (real_FFT_NS_Chart.Series[0].Points.Count > 0)
                    real_FFT_NS_Chart.Series[0].Points.Clear();

                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    if (f1[ii] >= 0.1)
                        real_FFT_NS_Chart.Series[0].Points.AddXY(f1[ii], FFTrealtime_signalNS[ii]);
                }

                //***********************************************************
                dataCounts = N;
                ss = new double[dataCounts];
                ss = signal3;
                #region FFT_EW
                //Calculating the Fourier Transform of Data
                //EW
                Y = new double[dataCounts][];
                imag = new double[dataCounts];

                Y = fft_dsa(ss, imag, dataCounts, 1);

                P2 = new double[dataCounts];
                for (int ii = 0; ii < dataCounts; ii++)
                {
                    P2[ii] = Math.Abs(Y[0][ii] / dataCounts);
                }
                P1 = new double[dataCounts / 2 + 1];

                for (int ii = 0; ii <= dataCounts / 2; ii++)
                {
                    P1[ii] = P2[ii];
                }
                for (int ii = 1; ii < dataCounts / 2; ii++)
                {
                    P1[ii] = 2 * P1[ii];
                }

                for (int ii = 0; ii < dataCounts / 2; ii++)
                {
                    FFTrealtime_signalEW[ii] = P1[ii];
                }
                #endregion

                f1 = new double[dataCounts / 2 + 1];
                for (int ii = 0; ii < dataCounts / 2 + 1; ii++)
                {
                    f1[ii] = Fs * ii / dataCounts;
                }

           if (real_FFT_EW_Chart.Series[0].Points.Count > 0)
                real_FFT_EW_Chart.Series[0].Points.Clear();
                for (int ii = 1; ii < dataCounts / 2 + 1; ii++)
                {
                    if (f1[ii] >= 0.1)
                        real_FFT_EW_Chart.Series[0].Points.AddXY(f1[ii], FFTrealtime_signalEW[ii]);
                }
            }
    }      
}
