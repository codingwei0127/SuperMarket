using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Data.SqlClient;
using System.IO;
using Modbus.Device;
using System.Collections;
using System.Windows.Forms.DataVisualization.Charting;
using System.Globalization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;

namespace modbus_rtu_sensor
{
    public partial class m : Form
    {
        string SQL_String = "";
        bool IsAllowLightToTurnOn = true; 
        bool IsAllowLightToTurnOff = true;

        int[] SetTime;
        int OpenHour = 0, OpenMin = 0, TurnOffHour = 0, TurnOffMin = 0;

        double H, G, I, PF, A, B, CCCCC;

        private void TodayPowerTimer_Tick(object sender, EventArgs e)
        {
            // 當日耗電曲線圖 -- 初始撈SQL資料表並且顯示
            String today = DateTime.Now.ToString("yyyy-MM-dd");

            var date = DateTime.Parse(today);
            var tomorrow_date = date.AddDays(1); // 現在時間加一天: DateTime

            // DateTime to string
            string date_str = date.ToString("yyyy/MM/dd HH:mm:ss");
            string tomorrow_date_str = tomorrow_date.ToString("yyyy/MM/dd HH:mm:ss");

            //查某一天的SQL資料
            ArrayList HistoryData = new ArrayList();
            HistoryData = ReadDayData(date_str, tomorrow_date_str);

            // 從ArrayList中取得第一個字典
            Dictionary<string, object> myDict;

            // 從陣列中取得第一個字典
            myDict = (Dictionary<string, object>)HistoryData[0];

            string[] Time = (string[])myDict["Time"];
            double[] Voltage = (double[])myDict["Voltage"];
            double[] Current = (double[])myDict["Current"];
            double[] Power = (double[])myDict["Power"];
            double[] PF = (double[])myDict["PF"];
            double[] RE0 = (double[])myDict["RE0"];
            double[] RE1 = (double[])myDict["RE1"];
            int ArrayLength = Time.Length;

            TodayChart.Series[0].Points.Clear();
            for (int i = 0; i < ArrayLength; i++)
            {
                DateTime parsedDate = DateTime.Parse(Time[i]);
                string dt_str = parsedDate.ToString("HH:mm");
                TodayChart.Series[0].Points.AddXY(dt_str, Power[i]);
            }
        } // TodayPowerTimer_Tick()

        private void RangeSearchButton_Click(object sender, EventArgs e)
        {        
            // 採集Calendar上所選擇的日期
            String startCalendar = StartCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");
            String endCalendar = EndCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");

            var startDate = DateTime.Parse(startCalendar);
            var endDate = DateTime.Parse(endCalendar);

            // DateTime to string
            string startDate_str = startDate.ToString("yyyy/MM/dd HH:mm:ss");
            string endDate_str = endDate.ToString("yyyy/MM/dd HH:mm:ss");


            //查某一天的SQL資料
            ArrayList HistoryData = new ArrayList();
            HistoryData = ReadDayData(startDate_str, endDate_str);

            // 從ArrayList中取得第一個字典
            Dictionary<string, object> myDict;

            // 從陣列中取得第一個字典
            myDict = (Dictionary<string, object>)HistoryData[0];

            // 從字典中取得時間, 以及其他量測資料
            string[] Time = (string[])myDict["Time"];
            double[] Voltage = (double[])myDict["Voltage"];
            double[] Current = (double[])myDict["Current"];
            double[] Power = (double[])myDict["Power"];
            double[] PF = (double[])myDict["PF"];
            double[] RE0 = (double[])myDict["RE0"];
            double[] RE1 = (double[])myDict["RE1"];
            int ArrayLength = Time.Length;

            if(ArrayLength == 0)
            {
                RangeSearchChart.Series[0].Points.Clear();
                // 讓chart消失
                RangeSearchChart.Visible = false;

                MessageBox.Show("查無資料!");
            }
            else
            {
                // 讓chart顯示
                RangeSearchChart.Visible = true;

                RangeSearchChart.Series[0].Points.Clear();
                for (int i = 0; i < ArrayLength; i++)
                {
                    RangeSearchChart.Series[0].Points.AddXY(Time[i], Power[i]);
                }
            }                   
        } // RangeSearchButton_Click()

        private void YearRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            YearComboBox.Visible = true;
            YearLabel.Visible = true;
            MonthLabel_month.Visible = false;
            MonthLabel_year.Visible = false;
            CalendarLabel.Visible = false;
            MonthComboBox_month.Visible = false;
            MonthComboBox_year.Visible = false;
            PeriodCalendar.Visible = false;
            PeriodSearchButton.Visible = false;
            ExcelPeriodButton.Visible = false;
        }

        private void MonthRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            MonthComboBox_month.Visible = true;
            MonthComboBox_year.Visible = true;
            MonthLabel_month.Visible = true;
            MonthLabel_year.Visible = true;
            YearLabel.Visible = false;
            CalendarLabel.Visible = false;
            PeriodCalendar.Visible = false;
            YearComboBox.Visible = false;
            PeriodSearchButton.Visible = false;
            ExcelPeriodButton.Visible = false;
        }

        private void DayRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            PeriodCalendar.Visible = true;
            CalendarLabel.Visible = true;
            MonthLabel_month.Visible = false;
            MonthLabel_year.Visible = false;
            YearLabel.Visible = false;
            MonthComboBox_month.Visible = false;
            MonthComboBox_year.Visible = false;
            YearComboBox.Visible = false;
            PeriodSearchButton.Visible = false;
            ExcelPeriodButton.Visible = false;
        }

        private void PeriodSearchButton_Click(object sender, EventArgs e)
        {
            if (DayRadioButton.Checked == true)
            {
                PeriodSearchChart.Visible = true;

                // 把曲線圖的形式改為Line
                PeriodSearchChart.Series[0].ChartType = SeriesChartType.Line;
                PeriodSearchChart.Series[0].LegendText = "耗電量(kW)";

                // 採集Calendar上所選擇的日期
                String myCalendar = PeriodCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");

                var date = DateTime.Parse(myCalendar);
                var tomorrow_date = date.AddDays(1); // 現在時間加一天: DateTime

                // DateTime to string
                string date_str = date.ToString("yyyy/MM/dd HH:mm:ss");
                string tomorrow_date_str = tomorrow_date.ToString("yyyy/MM/dd HH:mm:ss");

                //查某一天的SQL資料
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadDayData(date_str, tomorrow_date_str);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                int ArrayLength = Time.Length;

                if (ArrayLength == 0)
                {
                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    // 讓chart消失
                    PeriodSearchChart.Visible = false;

                    MessageBox.Show("查無資料!");
                }
                else
                {
                    PeriodSearchChart.Visible = true;

                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    for (int i = 0; i < ArrayLength; i++)
                    {
                        DateTime parsedDate = DateTime.Parse(Time[i]);
                        string dt_str = parsedDate.ToString("HH:mm");
                        PeriodSearchChart.Series[0].Points.AddXY(dt_str, Power[i]);
                    }
                }
            } // if (DayRadioButton.Checked == true)
            else if (MonthRadioButton.Checked == true) // 月查詢
            {

                PeriodSearchChart.Visible = true;

                // 把曲線圖的形式改為直條圖
                PeriodSearchChart.Series[0].ChartType = SeriesChartType.Column;
                PeriodSearchChart.Series[0].LegendText = "日平均總耗電量kWh(度)";

                string Year_str = (string)MonthComboBox_year.SelectedItem;
                string Month_str = (string)MonthComboBox_month.SelectedItem;

                //insert sql: READ_SQL
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadMonthData(Year_str, Month_str);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                int ArrayLength = Time.Length;

                if (ArrayLength == 0)
                {
                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    // 讓chart消失
                    PeriodSearchChart.Visible = false;

                    MessageBox.Show("查無資料!");
                }
                else
                {
                    PeriodSearchChart.Visible = true;

                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    //畫出曲線圖
                    for (int i = 0; i < ArrayLength; i++)
                    {
                        //把時間格式修改成 "月"
                        DateTime parsedDate = DateTime.Parse(Time[i]);
                        string day = parsedDate.Day + "號";

                        PeriodSearchChart.Series[0].Points.AddXY(day, Power[i]);
                    }
                }
            } // else if(MonthRadioButton.Checked == true)
            else if (YearRadioButton.Checked == true) // 年查詢
            {             
                // 把曲線圖的形式改為直條圖
                PeriodSearchChart.Series[0].ChartType = SeriesChartType.Column;
                PeriodSearchChart.Series[0].LegendText = "月平均總耗電量kWh(度)";

                Object selectedItem = YearComboBox.SelectedItem;
                string selectedItem_str = selectedItem.ToString();

                //insert sql: READ_SQL
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadYearData(selectedItem_str);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                int ArrayLength = Time.Length;

                if(ArrayLength == 0)
                {
                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    // 讓chart消失
                    PeriodSearchChart.Visible = false;

                    MessageBox.Show("查無資料!");
                }
                else
                {
                    PeriodSearchChart.Visible = true;

                    //清空曲線圖
                    PeriodSearchChart.Series[0].Points.Clear();

                    //畫出曲線圖
                    for (int i = 0; i < ArrayLength; i++)
                    {
                        //把時間格式修改成 "月"
                        DateTime parsedDate = DateTime.Parse(Time[i]);
                        string month = parsedDate.Month + "月";

                        PeriodSearchChart.Series[0].Points.AddXY(month, Power[i]);
                    }
                }   
            } // else if (YearRadioButton.Checked == true)
        }

        private void PeriodCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            PeriodSearchButton.Visible = true;
            ExcelPeriodButton.Visible = true;
        }

        private void YearComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            PeriodSearchButton.Visible = true;
            ExcelPeriodButton.Visible = true;
        }

        private void MonthComboBox_year_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MonthComboBox_month.SelectedIndex > -1) //somthing was selected
            {
                PeriodSearchButton.Visible = true;
                ExcelPeriodButton.Visible = true;
            }             
        }

        private void MonthComboBox_month_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MonthComboBox_year.SelectedIndex > -1) //somthing was selected
            {
                PeriodSearchButton.Visible = true;
                ExcelPeriodButton.Visible = true;
            }
        }

        public m()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) 
        {
            //撈初始資料: 已設定的開關燈時間
            SetTime = ReadSQL();
            OpenHour = SetTime[0];
            OpenMin = SetTime[1];
            TurnOffHour = SetTime[2];
            TurnOffMin = SetTime[3];

            SetOpenHour.Value = OpenHour;
            SetOpenMin.Value = OpenMin;
            SetTurnOffHour.Value = TurnOffHour;
            SetTurnOffMin.Value = TurnOffMin;

            ReadOpenHour.Text = Convert.ToString(OpenHour);
            ReadOpenMin.Text = Convert.ToString(OpenMin);
            ReadTurnOffHour.Text = Convert.ToString(TurnOffHour);
            ReadTurnOffMin.Text = Convert.ToString(TurnOffMin);

            //硬體設備的相關設定
            serialPort1.PortName = "COM4";     ////串列埠號
            serialPort1.BaudRate = 9600;          //鮑率
            serialPort1.DataBits = 8;             //資料位
            serialPort1.StopBits = System.IO.Ports.StopBits.One;
            serialPort1.Parity = System.IO.Ports.Parity.None;
            serialPort1.Open();
        } // Form1_Load()

        private void ExcelRangeButton_Click(object sender, EventArgs e)
        {
            // 採集Calendar上所選擇的日期
            String startCalendar = StartCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");
            String endCalendar = EndCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");

            var startDate = DateTime.Parse(startCalendar);
            var endDate = DateTime.Parse(endCalendar);

            // DateTime to string
            string startDate_str = startDate.ToString("yyyy/MM/dd HH:mm:ss");
            string endDate_str = endDate.ToString("yyyy/MM/dd HH:mm:ss");
            string excelRangeStr = startDate_str + "$" + endDate_str;

            //撈sql資料
            ArrayList HistoryData = new ArrayList();
            HistoryData = ReadExcelData("rangeQuery", excelRangeStr);

            // 從ArrayList中取得第一個字典
            Dictionary<string, object> myDict;

            // 從陣列中取得第一個字典
            myDict = (Dictionary<string, object>)HistoryData[0];

            // 從字典中取得時間, 以及其他量測資料
            string[] Time = (string[])myDict["Time"];
            double[] Voltage = (double[])myDict["Voltage"];
            double[] Current = (double[])myDict["Current"];
            double[] Power = (double[])myDict["Power"];
            double[] PF = (double[])myDict["PF"];
            double[] RE0 = (double[])myDict["RE0"];
            double[] RE1 = (double[])myDict["RE1"];
            double[] RE2 = (double[])myDict["RE2"];
            int ArrayLength = Time.Length;

            if(ArrayLength > 0)
            {
                // 創建 Excel Application 對象
                Application xlApp = new Application();

                // 如果 Excel 沒有安裝在系統中，則顯示錯誤信息
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                // 創建一個新的 Excel 文件
                Workbook xlWorkbook = xlApp.Workbooks.Add();
                // 創建一個新的工作表
                Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();

                // 在 Excel 工作表中創建標題行
                xlWorksheet.Cells[1, 1] = "時間";
                xlWorksheet.Cells[1, 2] = "電壓(V)";
                xlWorksheet.Cells[1, 3] = "電流(A)";
                xlWorksheet.Cells[1, 4] = "功率(kW)";
                xlWorksheet.Cells[1, 5] = "功率因數";
                xlWorksheet.Cells[1, 6] = "1號繼電器開關";
                xlWorksheet.Cells[1, 7] = "2號繼電器開關";
                xlWorksheet.Cells[1, 8] = "3號繼電器開關";

                // 在 Excel 工作表中寫入數據
                for (int i = 0; i < ArrayLength; i++)
                {
                    xlWorksheet.Cells[i + 2, 1] = Time[i];
                    xlWorksheet.Cells[i + 2, 2] = Voltage[i];
                    xlWorksheet.Cells[i + 2, 3] = Current[i];
                    xlWorksheet.Cells[i + 2, 4] = Power[i];
                    xlWorksheet.Cells[i + 2, 5] = PF[i];
                    xlWorksheet.Cells[i + 2, 6] = RE0[i];
                    xlWorksheet.Cells[i + 2, 7] = RE1[i];
                    xlWorksheet.Cells[i + 2, 8] = RE2[i];
                }

                // 保存 Excel 文件
                string fileName = "measurement.xlsx";
                xlWorkbook.SaveAs(fileName);

                // 關閉 Excel 文件和應用程序對象
                xlWorkbook.Close();
                xlApp.Quit();

                // 釋放 COM 對象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                // 顯示成功消息
                MessageBox.Show("Excel file created!!");
            }
            else
            {
                MessageBox.Show("查無資料可下載!");
            }
            
        } // ExcelRangeButton_Click()

        private void button1_Click(object sender, EventArgs e)
        {
            // 把設定的時間傳入SQL
            OpenHour = (int)SetOpenHour.Value;
            OpenMin = (int)SetOpenMin.Value;
            TurnOffHour = (int)SetTurnOffHour.Value;
            TurnOffMin = (int)SetTurnOffMin.Value;
            InsertSQL(OpenHour, OpenMin, TurnOffHour, TurnOffMin);

            // 設定的時間顯示在winforms的label上
            ReadOpenHour.Text = Convert.ToString((int)SetOpenHour.Value);
            ReadOpenMin.Text = Convert.ToString((int)SetOpenMin.Value);
            ReadTurnOffHour.Text = Convert.ToString((int)SetTurnOffHour.Value);
            ReadTurnOffMin.Text = Convert.ToString((int)SetTurnOffMin.Value);

            MessageBox.Show("設定成功!");
        }

        private void ExcelPeriodButton_Click(object sender, EventArgs e)
        {
            if (DayRadioButton.Checked == true)
            {
                // 採集Calendar上所選擇的日期
                String myCalendar = PeriodCalendar.SelectionRange.Start.ToString("yyyy-MM-dd");

                var date = DateTime.Parse(myCalendar);
                var tomorrow_date = date.AddDays(1); // 現在時間加一天: DateTime

                // DateTime to string
                string date_str = date.ToString("yyyy/MM/dd HH:mm:ss");
                string tomorrow_date_str = tomorrow_date.ToString("yyyy/MM/dd HH:mm:ss");

                string datetime = date_str + "$" + tomorrow_date_str;

                //撈sql資料
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadExcelData("Day", datetime);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                double[] RE2 = (double[])myDict["RE2"];
                int ArrayLength = Time.Length;

                if(ArrayLength > 0)
                {
                    // 創建 Excel Application 對象
                    Application xlApp = new Application();

                    // 如果 Excel 沒有安裝在系統中，則顯示錯誤信息
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }

                    // 創建一個新的 Excel 文件
                    Workbook xlWorkbook = xlApp.Workbooks.Add();
                    // 創建一個新的工作表
                    Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();

                    // 在 Excel 工作表中創建標題行
                    xlWorksheet.Cells[1, 1] = "時間";
                    xlWorksheet.Cells[1, 2] = "電壓(V)";
                    xlWorksheet.Cells[1, 3] = "電流(A)";
                    xlWorksheet.Cells[1, 4] = "功率(kW)";
                    xlWorksheet.Cells[1, 5] = "功率因數";
                    xlWorksheet.Cells[1, 6] = "1號繼電器開關";
                    xlWorksheet.Cells[1, 7] = "2號繼電器開關";
                    xlWorksheet.Cells[1, 8] = "3號繼電器開關";

                    // 在 Excel 工作表中寫入數據
                    for (int i = 0; i < ArrayLength; i++)
                    {
                        xlWorksheet.Cells[i + 2, 1] = Time[i];
                        xlWorksheet.Cells[i + 2, 2] = Voltage[i];
                        xlWorksheet.Cells[i + 2, 3] = Current[i];
                        xlWorksheet.Cells[i + 2, 4] = Power[i];
                        xlWorksheet.Cells[i + 2, 5] = PF[i];
                        xlWorksheet.Cells[i + 2, 6] = RE0[i];
                        xlWorksheet.Cells[i + 2, 7] = RE1[i];
                        xlWorksheet.Cells[i + 2, 8] = RE2[i];
                    }

                    // 保存 Excel 文件
                    string fileName = "measurement.xlsx";
                    xlWorkbook.SaveAs(fileName);

                    // 關閉 Excel 文件和應用程序對象
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // 釋放 COM 對象
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    // 顯示成功消息
                    MessageBox.Show("Excel file created!!");
                }
                else
                {
                    MessageBox.Show("查無資料可下載!");
                }

            } // if (DayRadioButton.Checked == true)
            else if (MonthRadioButton.Checked == true) // 月查詢
            {
                string Year_str = (string)MonthComboBox_year.SelectedItem;
                string Month_str = (string)MonthComboBox_month.SelectedItem;
                string datetime = Year_str + "$" + Month_str;

                //撈sql資料
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadExcelData("Month", datetime);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                double[] RE2 = (double[])myDict["RE2"];
                int ArrayLength = Time.Length;

                if(ArrayLength > 0)
                {
                    // 創建 Excel Application 對象
                    Application xlApp = new Application();

                    // 如果 Excel 沒有安裝在系統中，則顯示錯誤信息
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }

                    // 創建一個新的 Excel 文件
                    Workbook xlWorkbook = xlApp.Workbooks.Add();
                    // 創建一個新的工作表
                    Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();

                    // 在 Excel 工作表中創建標題行
                    xlWorksheet.Cells[1, 1] = "時間";
                    xlWorksheet.Cells[1, 2] = "電壓(V)";
                    xlWorksheet.Cells[1, 3] = "電流(A)";
                    xlWorksheet.Cells[1, 4] = "功率(kW)";
                    xlWorksheet.Cells[1, 5] = "功率因數";
                    xlWorksheet.Cells[1, 6] = "1號繼電器開關";
                    xlWorksheet.Cells[1, 7] = "2號繼電器開關";
                    xlWorksheet.Cells[1, 8] = "3號繼電器開關";

                    // 在 Excel 工作表中寫入數據
                    for (int i = 0; i < ArrayLength; i++)
                    {
                        xlWorksheet.Cells[i + 2, 1] = Time[i];
                        xlWorksheet.Cells[i + 2, 2] = Voltage[i];
                        xlWorksheet.Cells[i + 2, 3] = Current[i];
                        xlWorksheet.Cells[i + 2, 4] = Power[i];
                        xlWorksheet.Cells[i + 2, 5] = PF[i];
                        xlWorksheet.Cells[i + 2, 6] = RE0[i];
                        xlWorksheet.Cells[i + 2, 7] = RE1[i];
                        xlWorksheet.Cells[i + 2, 8] = RE2[i];

                    }

                    // 保存 Excel 文件
                    string fileName = "measurement.xlsx";
                    xlWorkbook.SaveAs(fileName);

                    // 關閉 Excel 文件和應用程序對象
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // 釋放 COM 對象
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    // 顯示成功消息
                    MessageBox.Show("Excel file created!!");
                }
                else
                {
                    MessageBox.Show("查無資料可下載!");
                }

            } // else if(MonthRadioButton.Checked == true)
            else if (YearRadioButton.Checked == true) // 年查詢
            {
                Object selectedItem = YearComboBox.SelectedItem;
                string selectedItem_str = selectedItem.ToString();

                //撈sql資料
                ArrayList HistoryData = new ArrayList();
                HistoryData = ReadExcelData("Year", selectedItem_str);

                // 從ArrayList中取得第一個字典
                Dictionary<string, object> myDict;

                // 從陣列中取得第一個字典
                myDict = (Dictionary<string, object>)HistoryData[0];

                // 從字典中取得時間, 以及其他量測資料
                string[] Time = (string[])myDict["Time"];
                double[] Voltage = (double[])myDict["Voltage"];
                double[] Current = (double[])myDict["Current"];
                double[] Power = (double[])myDict["Power"];
                double[] PF = (double[])myDict["PF"];
                double[] RE0 = (double[])myDict["RE0"];
                double[] RE1 = (double[])myDict["RE1"];
                double[] RE2 = (double[])myDict["RE2"];
                int ArrayLength = Time.Length;

                if(ArrayLength > 0)
                {
                    // 創建 Excel Application 對象
                    Application xlApp = new Application();

                    // 如果 Excel 沒有安裝在系統中，則顯示錯誤信息
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }

                    // 創建一個新的 Excel 文件
                    Workbook xlWorkbook = xlApp.Workbooks.Add();
                    // 創建一個新的工作表
                    Worksheet xlWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();

                    // 在 Excel 工作表中創建標題行
                    xlWorksheet.Cells[1, 1] = "時間";
                    xlWorksheet.Cells[1, 2] = "電壓(V)";
                    xlWorksheet.Cells[1, 3] = "電流(A)";
                    xlWorksheet.Cells[1, 4] = "功率(kW)";
                    xlWorksheet.Cells[1, 5] = "功率因數";
                    xlWorksheet.Cells[1, 6] = "1號繼電器開關";
                    xlWorksheet.Cells[1, 7] = "2號繼電器開關";
                    xlWorksheet.Cells[1, 8] = "3號繼電器開關";

                    // 在 Excel 工作表中寫入數據
                    for (int i = 0; i < ArrayLength; i++)
                    {
                        xlWorksheet.Cells[i + 2, 1] = Time[i];
                        xlWorksheet.Cells[i + 2, 2] = Voltage[i];
                        xlWorksheet.Cells[i + 2, 3] = Current[i];
                        xlWorksheet.Cells[i + 2, 4] = Power[i];
                        xlWorksheet.Cells[i + 2, 5] = PF[i];
                        xlWorksheet.Cells[i + 2, 6] = RE0[i];
                        xlWorksheet.Cells[i + 2, 7] = RE1[i];
                        xlWorksheet.Cells[i + 2, 8] = RE2[i];
                    }

                    // 保存 Excel 文件
                    string fileName = "measurement.xlsx";
                    xlWorkbook.SaveAs(fileName);

                    // 關閉 Excel 文件和應用程序對象
                    xlWorkbook.Close();
                    xlApp.Quit();

                    // 釋放 COM 對象
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    // 顯示成功消息
                    MessageBox.Show("Excel file created!!");
                }
                else
                {
                    MessageBox.Show("查無資料可下載!");
                }
            } // else if (YearRadioButton.Checked == true)                  
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            using (SqlConnection con01 = new SqlConnection(SQL_String))
            {
                ///open connection
                con01.Open();
                /////填入語法

                SqlCommand insert = new SqlCommand("INSERT INTO [MeterInformation].[dbo].[Information] ([Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1],[RE2]) VALUES (@value0,@value1,@value2,@value3,@value4,@value5,@value6,@value7)", con01);
                insert.Parameters.AddWithValue("@value0", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                insert.Parameters.AddWithValue("@value1", H);//電壓
                insert.Parameters.AddWithValue("@value2", G);//電流
                insert.Parameters.AddWithValue("@value3", I);//功率
                insert.Parameters.AddWithValue("@value4", PF);//PF
                insert.Parameters.AddWithValue("@value5", A);//繼電器1
                insert.Parameters.AddWithValue("@value6", B);//繼電器2
                insert.Parameters.AddWithValue("@value7", CCCCC);//繼電器3
                insert.ExecuteNonQuery();
            }
        }

        public void InsertSQL(int openHour, int openMin, int turnOffHour, int turnOffMin) // 把設定的\開關燈時間INSERT到SQL的資料表裡面
        {
            using (SqlConnection Conn01 = new SqlConnection(SQL_String))
            {
                SqlCommand insert = new SqlCommand("INSERT INTO [SuperMarketSchedule].[dbo].[Schedule]([Time] ,[OpenHour] ,[OpenMin], [TurnOffHour], [TurnOffMin])VALUES (@value0, @value1, @value2, @value3, @value4)", Conn01);
                Conn01.Open();// 開啟連線

                insert.Parameters.AddWithValue("@value0", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                insert.Parameters.AddWithValue("@value1", openHour);
                insert.Parameters.AddWithValue("@value2", openMin);
                insert.Parameters.AddWithValue("@value3", turnOffHour);
                insert.Parameters.AddWithValue("@value4", turnOffMin);

                insert.ExecuteNonQuery();//執行insert
                Conn01.Close();//斷開連線
            }
        }

        public int[] ReadSQL() //讀取排程時間
        {
            DataTable dtTable = new DataTable();
            SqlConnection Conn01 = new SqlConnection(SQL_String);
            SqlCommand cmd01 = new SqlCommand("SELECT TOP(1) [Time],[OpenHour],[OpenMin],[TurnOffHour],[TurnOffMin] FROM [SuperMarketSchedule].[dbo].[Schedule] ORDER BY [Time] DESC", Conn01);
            SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);

            Conn01.Open();
            bas1.Fill(dtTable);
            Conn01.Close();

            int Row = dtTable.Rows.Count; // sql資料表的列數

            if (Row > 0) // 確認SQL資料表中要有資料
            {
                OpenHour = Convert.ToInt32(dtTable.Rows[0][1]);
                OpenMin = Convert.ToInt32(dtTable.Rows[0][2]);
                TurnOffHour = Convert.ToInt32(dtTable.Rows[0][3]);
                TurnOffMin = Convert.ToInt32(dtTable.Rows[0][4]);
            }

            // 一維陣列
            int[] setTime = { OpenHour, OpenMin, TurnOffHour, TurnOffMin };
            return setTime;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            try
            {
                label1.Text = "時間: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");//時間顯示

                //開rs485通訊
                ModbusMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                master.Transport.ReadTimeout = 500;
                master.Transport.WriteTimeout = 500;
                master.Transport.Retries = 3;




                ushort[] holding_register1;
                holding_register1 = master.ReadHoldingRegisters(2, 5, 1);

                //H = Convert.ToDouble(holding_register1[0]) / 10;
                H = Math.Round(Convert.ToDouble(holding_register1[0]) / 10, 2, MidpointRounding.AwayFromZero);//電壓
                textBox1.Text = Convert.ToString(H);





                ushort[] holding_register;
                holding_register = master.ReadHoldingRegisters(2, 7, 1);//ID:2,位置:7,讀取筆數:1(電表的數據)

                //G = Math.Round(Convert.ToDouble(holding_register[0]) / 1000, 2);//電流
                G = Math.Round(Convert.ToDouble(holding_register[0]) / 1000, 2, MidpointRounding.AwayFromZero);//電流
                if(G < 10) // 土炮
                {
                    G += 65.5;
                }

                textBox2.Text = Convert.ToString(G);





                ushort[] holding_register2;
                holding_register2 = master.ReadHoldingRegisters(2, 17, 1);

                PF = Math.Round(Convert.ToDouble(holding_register2[0]) / 1000, 2, MidpointRounding.AwayFromZero);//功率因數
                textBox4.Text = Convert.ToString(PF);


                
                I = Math.Round(H * G * 1.732 * PF/1000, 2, MidpointRounding.AwayFromZero);
                textBox3.Text = Convert.ToString(I);


                bool[] holding_register3;
                holding_register3 = master.ReadCoils(128, 0, 1); //繼電器迴路1
                A = Convert.ToInt32(holding_register3[0]);//目前開關狀態
                textBox5.Text = Convert.ToString(A);


                bool[] holding_register4;
                holding_register4 = master.ReadCoils(128, 1, 1); //繼電器迴路2
                B = Convert.ToInt32(holding_register4[0]);//目前開關狀態
                textBox6.Text = Convert.ToString(B);

                

                string TimeNow = DateTime.Now.ToString();
                //TimeLabel.Text = TimeNow; // 讀取當前的時間

                int[] SetTime = ReadSQL();
                int OpenHour = 0, OpenMin = 0, TurnOffHour = 0, TurnOffMin = 0;
                OpenHour = SetTime[0];
                OpenMin = SetTime[1];
                TurnOffHour = SetTime[2];
                TurnOffMin = SetTime[3];

                // 開
                if (DateTime.Now.Hour == OpenHour && DateTime.Now.Minute == OpenMin && IsAllowLightToTurnOn)
                {
                    IsAllowLightToTurnOn = false; // 不要讓"開燈"這個動作一直重複
                    master.Transport.Retries = 0; //重讀機制
                    master.Transport.ReadTimeout = 10000; //作業逾時設定時間
                    master.WriteSingleCoil(128, 0, true); //繼電器站號，第幾顆繼電器(0，1，2，3)，繼電器開關
                    master.WriteSingleCoil(128, 1, true);
         
                    //MessageBox.Show("嘿嘿嘿開燈囉");
                }
                else if (DateTime.Now.Hour != OpenHour || DateTime.Now.Minute != OpenMin)
                {
                    IsAllowLightToTurnOn = true;
                }

                // Off
                if (DateTime.Now.Hour == TurnOffHour && DateTime.Now.Minute == TurnOffMin && IsAllowLightToTurnOff)
                {
                    IsAllowLightToTurnOff = false; // 不要讓Off這個動作一直重複
                    master.Transport.Retries = 1; //重讀機制
                    master.Transport.ReadTimeout = 1000; //作業逾時設定時間
                    master.WriteSingleCoil(128, 0, false); //繼電器/開關開
                    master.WriteSingleCoil(128, 1, false);
                    //MessageBox.Show("OFF!");
                }
                else if (DateTime.Now.Hour != TurnOffHour || DateTime.Now.Minute != TurnOffMin)
                {
                    IsAllowLightToTurnOff = true;
                }


               


                //抓每小時平均

                DataTable HourData = new DataTable();
                SqlConnection Conn01 = new SqlConnection(SQL_String);

                string sss = "SELECT Top(1) [Time],[Voltage],[Current],[Power] ,[PF]  FROM [MeterInformation].[dbo].[InformationHour] ORDER BY [Time] DESC";

                SqlCommand cmd01 = new SqlCommand(sss, Conn01);
                SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);


                Conn01.Open();
                bas1.Fill(HourData);
                Conn01.Close();


                int RowHourData_1 = HourData.Rows.Count;
                string Time = "";
                string Voltage = "";
                string Current = "";
                string Power = "";
                string PF11 = "";

                if (RowHourData_1 > 0)
                {
                    Time = HourData.Rows[0][0].ToString().Replace(" ", "");
                    Voltage = Convert.ToDouble(HourData.Rows[0][1]).ToString();
                    Current = Convert.ToDouble(HourData.Rows[0][2]).ToString();
                    Power = Convert.ToDouble(HourData.Rows[0][3]).ToString();
                    PF11 = Convert.ToDouble(HourData.Rows[0][4]).ToString();
                }

                textBox7.Text = Voltage;
                textBox8.Text = Current;
                textBox9.Text = Power;
                textBox10.Text = PF11;





                //抓每月平均

                DataTable MonthData = new DataTable();
                SqlConnection Conn02 = new SqlConnection(SQL_String);

                string sss1 = "SELECT Top(1) [Time],[Voltage],[Current],[Power] ,[PF]  FROM [MeterInformation].[dbo].[InformationMonth] ORDER BY [Time] DESC";

                SqlCommand cmd02 = new SqlCommand(sss1, Conn02);
                SqlDataAdapter bas2 = new SqlDataAdapter(cmd02);


                Conn02.Open();
                bas2.Fill(MonthData);
                Conn02.Close();


                int RowMonthData_2 = MonthData.Rows.Count;
                string Time1 = "";
                string Voltage1 = "";
                string Current1 = "";
                string Power1 = "";
                string PF12 = "";

                if (RowMonthData_2 > 0)
                {
                    Time1 = MonthData.Rows[0][0].ToString().Replace(" ", "");
                    Voltage1 = Convert.ToDouble(MonthData.Rows[0][1]).ToString();
                    Current1 = Convert.ToDouble(MonthData.Rows[0][2]).ToString();
                    Power1 = Convert.ToDouble(MonthData.Rows[0][3]).ToString();
                    PF12 = Convert.ToDouble(MonthData.Rows[0][4]).ToString();
                }

                textBox11.Text = Voltage1;
                textBox12.Text = Current1;
                textBox13.Text = Power1;
                textBox14.Text = PF12;





            }
            catch (Exception e1)//抓取錯誤資訊
            {
                MessageBox.Show(e1.Message);
            }

        }

        public ArrayList ReadYearData(string year) // 年查詢(時間正序排列)
        {
            DataTable dtTable = new DataTable();
            SqlConnection Conn01 = new SqlConnection(SQL_String);
            SqlCommand cmd01 = new SqlCommand("SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1] FROM [MeterInformation].[dbo].[InformationMonth] WHERE YEAR([Time]) = " + year + " ORDER BY [Time]", Conn01);
            SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);

            Conn01.Open();
            bas1.Fill(dtTable);
            Conn01.Close();

            int Row = dtTable.Rows.Count; // sql資料表的列數
            string[] Time = new string[Row];
            double[] Voltage = new double[Row];
            double[] Current = new double[Row];
            double[] Power = new double[Row];
            double[] PF = new double[Row];
            double[] RE0 = new double[Row];
            double[] RE1 = new double[Row];

            if (Row > 0)
            {
                for (int i = 0; i < Row; i++)
                {
                    Time[i] = dtTable.Rows[i][0].ToString().Replace(" ", "");
                    Voltage[i] = Convert.ToDouble(dtTable.Rows[i][1]);
                    Current[i] = Convert.ToDouble(dtTable.Rows[i][2]);
                    Power[i] = Convert.ToDouble(dtTable.Rows[i][3]);
                    PF[i] = Convert.ToDouble(dtTable.Rows[i][4]);
                    RE0[i] = Convert.ToDouble(dtTable.Rows[i][5]);
                    RE1[i] = Convert.ToDouble(dtTable.Rows[i][6]);
                }
            }

            //命名一個陣列清單叫做ReturnArray
            ArrayList ReturnArray = new ArrayList();
            //有一個字典叫做temp_dict
            Dictionary<string, object> temp_dict;
            //這本temp_dict字典是<string, object>的類型(前面是名字，後面是裝的東西的類型，所以這個代表前面是字串，後面是東西)
            temp_dict = new Dictionary<string, object>
            {
                { "Time", Time },
                { "Voltage", Voltage },
                { "Current", Current },
                { "Power", Power },
                { "PF", PF },
                { "RE0", RE0 },
                { "RE1", RE1 }
            };

            //ReturnArray這個陣列裡面.塞進(temp_dict這個東西)
            ReturnArray.Add(temp_dict);
            //傳回ReturnArray
            return ReturnArray;
        }
        public ArrayList ReadMonthData(string year, string month) // 月查詢(時間正序排列)
        {
            DataTable dtTable = new DataTable();
            SqlConnection Conn01 = new SqlConnection(SQL_String);
            SqlCommand cmd01 = new SqlCommand("SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1] FROM [MeterInformation].[dbo].[InformationDay] WHERE YEAR([Time]) = " + year + " AND MONTH([Time]) = " + month + " ORDER BY [Time]", Conn01);
            SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);

            Conn01.Open();
            bas1.Fill(dtTable);
            Conn01.Close();

            int Row = dtTable.Rows.Count; // sql資料表的列數
            string[] Time = new string[Row];
            double[] Voltage = new double[Row];
            double[] Current = new double[Row];
            double[] Power = new double[Row];
            double[] PF = new double[Row];
            double[] RE0 = new double[Row];
            double[] RE1 = new double[Row];

            if (Row > 0)
            {
                for (int i = 0; i < Row; i++)
                {
                    Time[i] = dtTable.Rows[i][0].ToString().Replace(" ", "");
                    Voltage[i] = Convert.ToDouble(dtTable.Rows[i][1]);
                    Current[i] = Convert.ToDouble(dtTable.Rows[i][2]);
                    Power[i] = Convert.ToDouble(dtTable.Rows[i][3]);
                    PF[i] = Convert.ToDouble(dtTable.Rows[i][4]);
                    RE0[i] = Convert.ToDouble(dtTable.Rows[i][5]);
                    RE1[i] = Convert.ToDouble(dtTable.Rows[i][6]);
                }
            }

            //命名一個陣列清單叫做ReturnArray
            ArrayList ReturnArray = new ArrayList();
            //有一個字典叫做temp_dict
            Dictionary<string, object> temp_dict;
            //這本temp_dict字典是<string, object>的類型(前面是名字，後面是裝的東西的類型，所以這個代表前面是字串，後面是東西)
            temp_dict = new Dictionary<string, object>
            {
                { "Time", Time },
                { "Voltage", Voltage },
                { "Current", Current },
                { "Power", Power },
                { "PF", PF },
                { "RE0", RE0 },
                { "RE1", RE1 }
            };

            //ReturnArray這個陣列裡面.塞進(temp_dict這個東西)
            ReturnArray.Add(temp_dict);
            //傳回ReturnArray
            return ReturnArray;
        }
        public ArrayList ReadDayData(string start, string end) // 日查詢(時間正序排列)
        {
            DataTable dtTable = new DataTable();
            SqlConnection Conn01 = new SqlConnection(SQL_String);
            SqlCommand cmd01 = new SqlCommand("SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1] FROM [MeterInformation].[dbo].[Information] WHERE [Time] BETWEEN '" + start + "' AND'" + end + "' ORDER BY [Time]", Conn01);
            SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);

            Conn01.Open();
            bas1.Fill(dtTable);
            Conn01.Close();

            int Row = dtTable.Rows.Count; // sql資料表的列數
            string[] Time = new string[Row];
            double[] Voltage = new double[Row];
            double[] Current = new double[Row];
            double[] Power = new double[Row];
            double[] PF = new double[Row];
            double[] RE0 = new double[Row];
            double[] RE1 = new double[Row];

            if (Row > 0)
            {
                for (int i = 0; i < Row; i++)
                {
                    Time[i] = dtTable.Rows[i][0].ToString().Replace(" ", "");
                    Voltage[i] = Convert.ToDouble(dtTable.Rows[i][1]);
                    Current[i] = Convert.ToDouble(dtTable.Rows[i][2]);
                    Power[i] = Convert.ToDouble(dtTable.Rows[i][3]);
                    PF[i] = Convert.ToDouble(dtTable.Rows[i][4]);
                    RE0[i] = Convert.ToDouble(dtTable.Rows[i][5]);
                    RE1[i] = Convert.ToDouble(dtTable.Rows[i][6]);
                }
            }

            //命名一個陣列清單叫做ReturnArray
            ArrayList ReturnArray = new ArrayList();
            //有一個字典叫做temp_dict
            Dictionary<string, object> temp_dict;
            //這本temp_dict字典是<string, object>的類型(前面是名字，後面是裝的東西的類型，所以這個代表前面是字串，後面是東西)
            temp_dict = new Dictionary<string, object>
            {
                { "Time", Time },
                { "Voltage", Voltage },
                { "Current", Current },
                { "Power", Power },
                { "PF", PF },
                { "RE0", RE0 },
                { "RE1", RE1 }
            };

            //ReturnArray這個陣列裡面.塞進(temp_dict這個東西)
            ReturnArray.Add(temp_dict);
            //傳回ReturnArray
            return ReturnArray;
        }
        public ArrayList ReadExcelData(string queryMode, string time) // Excel讀SQL(時間正序排列)
        {

            string sqlConnection = "";
            if (queryMode == "Year")
            {
                sqlConnection = "SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1],[RE2] FROM [MeterInformation].[dbo].[Information] WHERE YEAR([Time]) = " + time + " ORDER BY [Time]";
            }
            else if (queryMode == "Month")
            {
                string[] datetimes = time.Split('$');
                sqlConnection = "SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1],[RE2] FROM [MeterInformation].[dbo].[Information] WHERE YEAR([Time]) = " + datetimes[0] + " AND MONTH([Time]) = " + datetimes[1] + " ORDER BY [Time]";
            }
            else if (queryMode == "Day" || queryMode == "rangeQuery")
            {
                // 2023-02-25, 2023-02-26
                string[] datetimes = time.Split('$');
                sqlConnection = "SELECT [Time],[Voltage],[Current],[Power],[PF],[RE0],[RE1],[RE2] FROM [MeterInformation].[dbo].[Information] WHERE [Time] BETWEEN '" + datetimes[0] + "' AND '" + datetimes[1] + "' ORDER BY [Time]";         
            }

            DataTable dtTable = new DataTable();
            SqlConnection Conn01 = new SqlConnection(SQL_String);
            SqlCommand cmd01 = new SqlCommand(sqlConnection, Conn01);
            SqlDataAdapter bas1 = new SqlDataAdapter(cmd01);

            Conn01.Open();
            bas1.Fill(dtTable);
            Conn01.Close();

            int Row = dtTable.Rows.Count; // sql資料表的列數
            string[] Time = new string[Row];
            double[] Voltage = new double[Row];
            double[] Current = new double[Row];
            double[] Power = new double[Row];
            double[] PF = new double[Row];
            double[] RE0 = new double[Row];
            double[] RE1 = new double[Row];
            double[] RE2 = new double[Row];

            if (Row > 0)
            {
                for (int i = 0; i < Row; i++)
                {
                    Time[i] = dtTable.Rows[i][0].ToString().Replace(" ", "");
                    Voltage[i] = Convert.ToDouble(dtTable.Rows[i][1]);
                    Current[i] = Convert.ToDouble(dtTable.Rows[i][2]);
                    Power[i] = Convert.ToDouble(dtTable.Rows[i][3]);
                    PF[i] = Convert.ToDouble(dtTable.Rows[i][4]);
                    RE0[i] = Convert.ToDouble(dtTable.Rows[i][5]);
                    RE1[i] = Convert.ToDouble(dtTable.Rows[i][6]);
                    RE2[i] = Convert.ToDouble(dtTable.Rows[i][7]);
                }
            }

            //命名一個陣列清單叫做ReturnArray
            ArrayList ReturnArray = new ArrayList();
            //有一個字典叫做temp_dict
            Dictionary<string, object> temp_dict;
            //這本temp_dict字典是<string, object>的類型(前面是名字，後面是裝的東西的類型，所以這個代表前面是字串，後面是東西)
            temp_dict = new Dictionary<string, object>
            {
                { "Time", Time },
                { "Voltage", Voltage },
                { "Current", Current },
                { "Power", Power },
                { "PF", PF },
                { "RE0", RE0 },
                { "RE1", RE1 },
                { "RE2", RE2 },
                { "Row", Row }
            };

            //ReturnArray這個陣列裡面.塞進(temp_dict這個東西)
            ReturnArray.Add(temp_dict);
            //傳回ReturnArray
            return ReturnArray;
        }

    }
}
