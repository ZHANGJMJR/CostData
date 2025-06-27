using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Globalization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Drawing;
using ZstdSharp.Unsafe;

namespace CostData
{
    public partial class Ribbon1
    {
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker; 
        private RibbonButton myButton;
        private string connectionString = "Server=127.0.0.1;Port=3309;Database=hotel;User ID=root;Password=123456;";

        private string getRegStr()
        {
            string keyPath = @"SOFTWARE\WOW6432Node\兴业银锡\XYAPP";
            string keyName = "IP"; // 你想要获取值的键名

            try
            {
                // 打开HKEY_LOCAL_MACHINE下的指定路径
                using (RegistryKey localMachineKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                {
                    using (RegistryKey key = localMachineKey.OpenSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            // 读取值
                            object value = key.GetValue(keyName);
                            if (value != null)
                            {
                                return value.ToString();
                            }
                            else
                            {
                                return "127.0.0.1:3309";
                            }
                        }
                        else
                        {
                            return "127.0.0.1:3309";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "127.0.0.1:3309";
            }
        }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            startDatePicker = new DateTimePicker();
            endDatePicker = new DateTimePicker(); 

            startDatePicker.Format = DateTimePickerFormat.Short;
            endDatePicker.Format = DateTimePickerFormat.Short;

            // 设置控件位置和大小
            startDatePicker.Width = 100;
            endDatePicker.Width = 100;
           // downloadDataButton.Text = "下载数据";

            // 将控件添加到 Ribbon
            //this.tab1.Groups[0].Items.Add(startDatePicker);
            //this.tab1.Groups[0].Items.Add(endDatePicker);
            //this.tab1.Groups[0].Items.Add(downloadDataButton);

            // 绑定按钮点击事件
            //this.tab1.Groups[0].button1.Click += DownloadDataButton_Click;
           // myButton = this.tab1.Groups[0].("button1") as RibbonButton;

        }
        //private void DownloadDataButton_Click(object sender, EventArgs e)
        //{
        //    DateTime startDate = startDatePicker.Value;
        //    DateTime endDate = endDatePicker.Value;
        //    string lsip = getRegStr();


        //    using (MySqlConnection conn = new MySqlConnection(connectionString))
        //    {
        //        try
        //        {
        //            conn.Open();
        //            string query = "SELECT * FROM your_table WHERE datadt BETWEEN @startDate AND @endDate";
        //            MySqlCommand cmd = new MySqlCommand(query, conn);
        //            cmd.Parameters.AddWithValue("@startDate", startDate);
        //            cmd.Parameters.AddWithValue("@endDate", endDate);

        //            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
        //            DataTable dataTable = new DataTable();
        //            adapter.Fill(dataTable);

        //            // 将数据填充到 Excel 工作表
        //            Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
        //            for (int i = 0; i < dataTable.Rows.Count; i++)
        //            {
        //                for (int j = 0; j < dataTable.Columns.Count; j++)
        //                {
        //                    worksheet.Cells[i + 1, j + 1] = dataTable.Rows[i][j].ToString();
        //                }
        //            }
        //            MessageBox.Show("数据下载成功！");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"错误: {ex.Message}");
        //        }
        //    }
        //}

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // 

        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void editBox3_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        public static bool IsValidDate(string dateString, string format)
        {
            DateTime dateTime;
            return DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);
        }
        private void download_btn_Click(object sender, RibbonControlEventArgs e)
        {
            string ls_start_date, ls_end_date;
            ls_start_date = this.editBox1.Text;
            ls_end_date = this.editBox2.Text;
            DateTime start_date, end_date;
            bool is_showmsg = false;
            string[] lsip = getRegStr().Split(':');
            if (!lsip[1].Equals("3309")) {
                connectionString = $"Server={lsip[0]};Port={lsip[1]};Database=hotel;User ID=root;Password=arzfUh??p3<L;";
            }
            if (!IsValidDate(ls_start_date, "yyyy-MM-dd"))
            {
                MessageBox.Show(ls_start_date + " 开始日期无效，请检查!");
                return;
            }
            start_date   = DateTime.Parse(ls_start_date);
            if (!IsValidDate(ls_end_date, "yyyy-MM-dd"))
            {
                MessageBox.Show(ls_end_date + " 结束日期无效，请检查!"); 
                return;
            }
            end_date = DateTime.Parse(ls_end_date);
            if (string.Compare(ls_start_date, ls_end_date)>0)
            {
                MessageBox.Show(  "开始日期  大于  结束日期无效，请检查!");
                return;
            }
            Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            int max_row = 1;
            max_row = worksheet.UsedRange.Rows.Count;
            max_row = max_row > 1 ? max_row + 1 : max_row;
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                try
                {
                    for (DateTime currentDate = start_date; currentDate <= end_date; currentDate = currentDate.AddDays(1))
                    {
                        string ls_worksheetname = worksheet.Name;
                        string ls_sh_worksheetname = currentDate.ToString("M月"); 

                        Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                        Worksheet targetSheet = null;
                        if (activeWorkbook != null)
                        {
                            if (activeWorkbook.Sheets.Count > 0) {
                                foreach (Worksheet sheet in activeWorkbook.Sheets)
                                {
                                    if (sheet.Name == ls_sh_worksheetname)
                                    {
                                        max_row = sheet.UsedRange.Rows.Count;
                                        max_row = max_row > 1 ? max_row + 1 : max_row;
                                        targetSheet = sheet;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                targetSheet = activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[activeWorkbook.Sheets.Count]);
                                targetSheet.Name = ls_sh_worksheetname;
                                max_row = 1;
                            }
                            if (targetSheet == null)
                            {
                                if (max_row == 1)
                                {
                                    if (worksheet.Name != ls_sh_worksheetname)
                                    {
                                        targetSheet = worksheet;
                                        targetSheet.Name = ls_sh_worksheetname;
                                    }
                                }
                                else
                                {
                                    if (worksheet.Name == ls_sh_worksheetname)
                                    {
                                        targetSheet = worksheet;
                                    }
                                    else
                                    {
                                        targetSheet = activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[activeWorkbook.Sheets.Count]);
                                        targetSheet.Name = ls_sh_worksheetname;
                                        max_row = 1;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("EXCEL异常，无法继续！");
                            return;
                        }
                        worksheet = targetSheet;
                        if (worksheet == null) {
                            MessageBox.Show("EXCEL异常，无法继续！");
                            return;
                        }
                        worksheet.Activate();
                        worksheet.Range["A2"].Select();
                        worksheet.Application.ActiveWindow.FreezePanes = true;

                        if (worksheet.Cells[1, "A"].Value2?.ToString() != "日期"|| worksheet.Cells[1, "B"].Value2?.ToString() != "Check Number" )
                        {
                            Excel.Range headerRange = worksheet.Range["A1", "Q1"];

                            // 设置第一行 A 至 H 列的背景色为 #FFE699
                            // 设置第一行 A 至 H 列的背景颜色为 #FFE699
                            Excel.Range rangeAH = worksheet.Range["A1:H1"];
                            rangeAH.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#FFE699"));

                            // 设置第一行 I 至 Q 列的背景颜色为 #C6E0B4
                            Excel.Range rangeIQ = worksheet.Range["I1:Q1"];
                            rangeIQ.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#C6E0B4"));

                            // 设置字体加粗
                            headerRange.Font.Bold = true;

                            // 设置字体颜色
                            headerRange.Font.Color = Excel.XlRgbColor.rgbBlack;

                            // 设置水平对齐方式为居中
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            // 设置垂直对齐方式为居中
                            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                            // 可选：自动调整列宽
                            headerRange.Columns.AutoFit();
                            string[] headers = {
                                "日期", "Check Number", "Open Time", "Check Total", "净价", "人数", "餐厅","",
                                "日期", "订单号", "订单金额", "实付金额", "净价", "订单状态", "支付类型", "完成时间"
                            };

                            // 填充表头数据到 A1:Q1
                            for (int i = 0; i < headers.Length; i++)
                            {
                                worksheet.Cells[1, i + 1] = headers[i];
                            }
                            worksheet.Range["B:E"].ColumnWidth = 15;
                            worksheet.Range["C:C"].ColumnWidth = 18;
                            worksheet.Range["G:G"].ColumnWidth = 12;
                            worksheet.Range["Q:Q"].ColumnWidth = 12;
                            worksheet.Range["J:P"].ColumnWidth = 15;
                           
                           // max_row = worksheet.UsedRange.Rows.Count + 1;
                        }
                        if (worksheet.Name != ls_sh_worksheetname)
                        { worksheet.Name = ls_sh_worksheetname;  }
                            max_row = worksheet.UsedRange.Rows.Count + 1;                       
                            List<Tmpdata> dataList = ReceiveData.FetchTmpdataByDate(conn, currentDate.ToString("yyyy-MM-dd"));
                        int inital_rows;
                        inital_rows = max_row;
                        if (dataList.Count<=0)
                        {
                            MessageBox.Show($"无{currentDate.ToString("yyyy-MM-dd")}日数据！");
                            continue;
                        }
                        for (int i = 0; i < dataList.Count; i++)
                        {
                            worksheet.Cells[i + max_row, "B"] = dataList[i].CheckNum.ToString()??"";
                            worksheet.Cells[i + max_row, "C"] = dataList[i].OpenDateTime.ToString() ?? "";
                            worksheet.Cells[i + max_row, "D"] = dataList[i].CheckTotal.ToString() ?? "";
                            worksheet.Cells[i + max_row, "E"] = $@"=IF(H{i + max_row} = """", round(D{i + max_row} / 1.06,2), IF(ISNUMBER(H{i + max_row}), round(D{i + max_row} / (100 + H{i + max_row})*100 ,2), 0))";
                            worksheet.Cells[i + max_row, "F"] = dataList[i].Guestnum.ToString()??"";
                            worksheet.Cells[i + max_row, "G"] = dataList[i].Rcsname?? "";
                            worksheet.Cells[i + max_row, "J"] = dataList[i].FZGOrderNumber.ToString()??"";
                            worksheet.Cells[i + max_row, "K"] = dataList[i].FZGOrderAmount.ToString()??"";
                            worksheet.Cells[i + max_row, "L"] = dataList[i].FZGReceivedAmount.ToString();
                            worksheet.Cells[i + max_row, "M"] = $@"=IF(H{i + max_row} = """",round( L{i + max_row} / 1.06,2), IF(ISNUMBER(H{i + max_row}), round(L{i + max_row} / (100 + H{i + max_row})*100,2) , 0))";
                            worksheet.Cells[i + max_row, "N"] = dataList[i].FZGOrderStatus??"";
                            worksheet.Cells[i + max_row, "O"] = dataList[i].FZGPayType??"";
                            worksheet.Cells[i + max_row, "P"] = dataList[i].FZGCompletionTime??"";
                            if (dataList[i].FirstName!=null && 
                                dataList[i].FirstName.ToUpper()=="ADD" && 
                                dataList[i].Daypart == "早")
                            {
                                worksheet.Cells[i + max_row, "Q"] = "慕味早餐";
                            }
                            else
                            {
                                worksheet.Cells[i + max_row, "Q"] = dataList[i].Rcsname;
                            }   
                        }
                        is_showmsg = true;
                        worksheet.Cells[dataList.Count + max_row, "D"] = $"= SUBTOTAL(9, D{inital_rows}: D{dataList.Count + max_row-1})";
                        worksheet.Cells[dataList.Count + max_row, "L"] = $"= SUBTOTAL(9, L{inital_rows}: L{dataList.Count + max_row-1})";
                        worksheet.Cells[dataList.Count + max_row, "D"].Font.Bold = true;
                        worksheet.Cells[dataList.Count + max_row, "L"].Font.Bold = true;
                        Excel.Range range = worksheet.Range[$"A{dataList.Count + max_row}:C{dataList.Count + max_row}"];
                        range.Merge(Type.Missing);  
                        range.Value2 = "汇总";                           
                        range.Font.Bold = true;                         
                        range.Font.Size = 14;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        worksheet.Range[$"A{dataList.Count + max_row}:Q{dataList.Count + max_row}"].
                            Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E7E6E6"));
                        worksheet.Cells[max_row, 1].Select();
                        max_row += dataList.Count+1;
                        worksheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        worksheet.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }
                    if(is_showmsg) MessageBox.Show("数据下载成功！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"错误: {ex.Message}");
                }
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void ApplyBanquetHeaderToActiveSheet(Excel.Worksheet targetSheet,int initalrow=4,string strmonth="")
        {

            int cellcolor = 14150647;
            int initrow = initalrow-3;
            // Cell [1, 1]
            {
                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 22], targetSheet.Cells[initalrow + 1, 22]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 22], targetSheet.Cells[initalrow + 1, 22]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 22], targetSheet.Cells[initalrow + 1, 22]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 23], targetSheet.Cells[initalrow + 1, 23]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 23], targetSheet.Cells[initalrow + 1, 23]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 23], targetSheet.Cells[initalrow + 1, 23]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 24], targetSheet.Cells[initalrow + 1, 24]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 24], targetSheet.Cells[initalrow + 1, 24]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 24], targetSheet.Cells[initalrow + 1, 24]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                targetSheet.Range[targetSheet.Cells[initalrow, 25], targetSheet.Cells[initalrow + 1, 25]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 25], targetSheet.Cells[initalrow + 1, 25]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 25], targetSheet.Cells[initalrow + 1, 25]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;


                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 11], targetSheet.Cells[initalrow, 14]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 11], targetSheet.Cells[initalrow, 14]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 11], targetSheet.Cells[initalrow, 14]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow + 1, 10]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow + 1, 10]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow + 1, 10]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Interior.Color = cellcolor;
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                // Cell [1, 1]
                targetSheet.Cells[initalrow, 1].Value2 = "Date";
                targetSheet.Cells[initalrow, 1].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 1].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 1].Font.Bold = true;
                targetSheet.Cells[initalrow, 1].Font.Italic = false;
                targetSheet.Cells[initalrow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 1].WrapText = true;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 2]
                targetSheet.Cells[initalrow, 2].Value2 = "Check No";
                targetSheet.Cells[initalrow, 2].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 2].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 2].Font.Bold = true;
                targetSheet.Cells[initalrow, 2].Font.Italic = false;
                targetSheet.Cells[initalrow, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 2].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 2].WrapText = true;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 3]
                targetSheet.Cells[initalrow, 3].Value2 = "Account Name";
                targetSheet.Cells[initalrow, 3].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 3].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 3].Font.Bold = true;
                targetSheet.Cells[initalrow, 3].Font.Italic = false;
                targetSheet.Cells[initalrow, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 3].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 3].WrapText = true;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 4]
                targetSheet.Cells[initalrow, 4].Value2 = "预订人";
                targetSheet.Cells[initalrow, 4].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 4].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 4].Font.Bold = true;
                targetSheet.Cells[initalrow, 4].Font.Italic = false;
                targetSheet.Cells[initalrow, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 4].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 4].WrapText = true;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 5]
                targetSheet.Cells[initalrow, 5].Value2 = "场地";
                targetSheet.Cells[initalrow, 5].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 5].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 5].Font.Bold = true;
                targetSheet.Cells[initalrow, 5].Font.Italic = false;
                targetSheet.Cells[initalrow, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 5].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 5].WrapText = true;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 6]
                targetSheet.Cells[initalrow, 6].Value2 = "Segments";
                targetSheet.Cells[initalrow, 6].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 6].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 6].Font.Bold = true;
                targetSheet.Cells[initalrow, 6].Font.Italic = false;
                targetSheet.Cells[initalrow, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 6].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 6].WrapText = true;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 7]
                targetSheet.Cells[initalrow, 7].Value2 = "Covers";
                targetSheet.Cells[initalrow, 7].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 7].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 7].Font.Bold = true;
                targetSheet.Cells[initalrow, 7].Font.Italic = false;
                targetSheet.Cells[initalrow, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 7].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 7].WrapText = true;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 8]
                targetSheet.Cells[initalrow, 8].Value2 = "Table";
                targetSheet.Cells[initalrow, 8].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 8].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 8].Font.Bold = true;
                targetSheet.Cells[initalrow, 8].Font.Italic = false;
                targetSheet.Cells[initalrow, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 8].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 8].WrapText = true;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 9]
                targetSheet.Cells[initalrow, 9].Value2 = "Per";
                targetSheet.Cells[initalrow, 9].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 9].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 9].Font.Bold = true;
                targetSheet.Cells[initalrow, 9].Font.Italic = false;
                targetSheet.Cells[initalrow, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 9].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 9].WrapText = true;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 10]
                targetSheet.Cells[initalrow, 10].Value2 = "%";
                targetSheet.Cells[initalrow, 10].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 10].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 10].Font.Bold = true;
                targetSheet.Cells[initalrow, 10].Font.Italic = false;
                targetSheet.Cells[initalrow, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 10].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 10].WrapText = true;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 11]
                targetSheet.Cells[initalrow, 11].Value2 = "Food";
                targetSheet.Cells[initalrow, 11].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 11].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 11].Font.Bold = true;
                targetSheet.Cells[initalrow, 11].Font.Italic = false;
                targetSheet.Cells[initalrow, 11].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 11].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 11].WrapText = true;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 12]
                targetSheet.Cells[initalrow, 12].Value2 = "";
                targetSheet.Cells[initalrow, 12].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 12].Font.Size = 12.0;
                targetSheet.Cells[initalrow, 12].Font.Bold = false;
                targetSheet.Cells[initalrow, 12].Font.Italic = false;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 13]
                targetSheet.Cells[initalrow, 13].Value2 = "";
                targetSheet.Cells[initalrow, 13].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 13].Font.Size = 12.0;
                targetSheet.Cells[initalrow, 13].Font.Bold = false;
                targetSheet.Cells[initalrow, 13].Font.Italic = false;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 14]
                targetSheet.Cells[initalrow, 14].Value2 = "";
                targetSheet.Cells[initalrow, 14].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 14].Font.Size = 12.0;
                targetSheet.Cells[initalrow, 14].Font.Bold = false;
                targetSheet.Cells[initalrow, 14].Font.Italic = false;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 15]
                targetSheet.Cells[initalrow, 15].Value2 = "Bev";
                targetSheet.Cells[initalrow, 15].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 15].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 15].Font.Bold = true;
                targetSheet.Cells[initalrow, 15].Font.Italic = false;
                targetSheet.Cells[initalrow, 15].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 15].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 15].WrapText = true;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 16]
                targetSheet.Cells[initalrow, 16].Value2 = "杂项";
                targetSheet.Cells[initalrow, 16].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 16].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 16].Font.Bold = true;
                targetSheet.Cells[initalrow, 16].Font.Italic = false;
                targetSheet.Cells[initalrow, 16].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 16].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 16].WrapText = true;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 17]
                targetSheet.Cells[initalrow, 17].Value2 = "啤酒";
                targetSheet.Cells[initalrow, 17].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 17].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 17].Font.Bold = true;
                targetSheet.Cells[initalrow, 17].Font.Italic = false;
                targetSheet.Cells[initalrow, 17].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 17].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 17].WrapText = true;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 18]
                targetSheet.Cells[initalrow, 18].Value2 = "能源";
                targetSheet.Cells[initalrow, 18].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 18].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 18].Font.Bold = true;
                targetSheet.Cells[initalrow, 18].Font.Italic = false;
                targetSheet.Cells[initalrow, 18].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 18].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 18].WrapText = true;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 19]
                targetSheet.Cells[initalrow, 19].Value2 = "场租";
                targetSheet.Cells[initalrow, 19].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 19].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 19].Font.Bold = true;
                targetSheet.Cells[initalrow, 19].Font.Italic = false;
                targetSheet.Cells[initalrow, 19].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 19].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 19].WrapText = true;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 20]
                targetSheet.Cells[initalrow, 20].Value2 = "Cigar";
                targetSheet.Cells[initalrow, 20].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 20].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 20].Font.Bold = true;
                targetSheet.Cells[initalrow, 20].Font.Italic = false;
                targetSheet.Cells[initalrow, 20].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 20].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 20].WrapText = true;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 21]
                targetSheet.Cells[initalrow, 21].Value2 = "Tax";
                targetSheet.Cells[initalrow, 21].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 21].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 21].Font.Bold = true;
                targetSheet.Cells[initalrow, 21].Font.Italic = false;
                targetSheet.Cells[initalrow, 21].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 21].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 21].WrapText = true;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 22]
                targetSheet.Cells[initalrow, 22].Value2 = "Total";
                targetSheet.Cells[initalrow, 22].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 22].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 22].Font.Bold = true;
                targetSheet.Cells[initalrow, 22].Font.Italic = false;
                targetSheet.Cells[initalrow, 22].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 22].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 22].WrapText = true;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 22].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 23]
                targetSheet.Cells[initalrow, 23].Value2 = "净价";
                targetSheet.Cells[initalrow, 23].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 23].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 23].Font.Bold = true;
                targetSheet.Cells[initalrow, 23].Font.Italic = false;
                targetSheet.Cells[initalrow, 23].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 23].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 23].WrapText = true;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 23].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 24]
                targetSheet.Cells[initalrow, 24].Value2 = "结账方式";
                targetSheet.Cells[initalrow, 24].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 24].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 24].Font.Bold = true;
                targetSheet.Cells[initalrow, 24].Font.Italic = false;
                targetSheet.Cells[initalrow, 24].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 24].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 24].WrapText = true;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 24].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 25]
                targetSheet.Cells[initalrow, 25].Value2 = "CheckTotal";
                targetSheet.Cells[initalrow, 25].Font.Name = "Arial";
                targetSheet.Cells[initalrow, 25].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 25].Font.Bold = true;
                targetSheet.Cells[initalrow, 25].Font.Italic = false;
                targetSheet.Cells[initalrow, 25].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 25].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 25].WrapText = true;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 25].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 1]
                targetSheet.Cells[initalrow + 1, 1].Value2 = "";
                targetSheet.Cells[initalrow + 1, 1].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 1].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 1].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 1].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 2]
                targetSheet.Cells[initalrow + 1, 2].Value2 = "";
                targetSheet.Cells[initalrow + 1, 2].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 2].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 2].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 2].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 3]
                targetSheet.Cells[initalrow + 1, 3].Value2 = "";
                targetSheet.Cells[initalrow + 1, 3].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 3].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 3].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 3].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 4]
                targetSheet.Cells[initalrow + 1, 4].Value2 = "";
                targetSheet.Cells[initalrow + 1, 4].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 4].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 4].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 4].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 5]
                targetSheet.Cells[initalrow + 1, 5].Value2 = "";
                targetSheet.Cells[initalrow + 1, 5].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 5].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 5].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 5].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 6]
                targetSheet.Cells[initalrow + 1, 6].Value2 = "";
                targetSheet.Cells[initalrow + 1, 6].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 6].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 6].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 6].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 7]
                targetSheet.Cells[initalrow + 1, 7].Value2 = "";
                targetSheet.Cells[initalrow + 1, 7].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 7].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 7].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 7].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 8]
                targetSheet.Cells[initalrow + 1, 8].Value2 = "";
                targetSheet.Cells[initalrow + 1, 8].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 8].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 8].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 8].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 9]
                targetSheet.Cells[initalrow + 1, 9].Value2 = "";
                targetSheet.Cells[initalrow + 1, 9].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 9].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 9].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 9].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 10]
                targetSheet.Cells[initalrow + 1, 10].Value2 = "";
                targetSheet.Cells[initalrow + 1, 10].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 10].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 10].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 10].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 11]
                targetSheet.Cells[initalrow + 1, 11].Value2 = "Breakfast";
                targetSheet.Cells[initalrow + 1, 11].Font.Name = "Arial";
                targetSheet.Cells[initalrow + 1, 11].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 11].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 11].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 11].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 11].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 11].WrapText = true;
                targetSheet.Cells[initalrow + 1, 11].Interior.Color = cellcolor;
                targetSheet.Cells[initalrow + 1, 11].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 12]
                targetSheet.Cells[initalrow + 1, 12].Value2 = "Lunch";
                targetSheet.Cells[initalrow + 1, 12].Font.Name = "Arial";
                targetSheet.Cells[initalrow + 1, 12].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 12].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 12].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 12].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 12].WrapText = true;
                targetSheet.Cells[initalrow + 1, 12].Interior.Color = cellcolor;
                targetSheet.Cells[initalrow + 1, 12].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 13]
                targetSheet.Cells[initalrow + 1, 13].Value2 = "Dinner";
                targetSheet.Cells[initalrow + 1, 13].Font.Name = "Arial";
                targetSheet.Cells[initalrow + 1, 13].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 13].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 13].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 13].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 13].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 13].WrapText = true;
                targetSheet.Cells[initalrow + 1, 13].Interior.Color = cellcolor;
                targetSheet.Cells[initalrow + 1, 13].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 14]
                targetSheet.Cells[initalrow + 1, 14].Value2 = "茶歇";
                targetSheet.Cells[initalrow + 1, 14].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 14].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 14].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 14].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 14].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 14].WrapText = true;
                targetSheet.Cells[initalrow + 1, 14].Interior.Color = cellcolor;
                targetSheet.Cells[initalrow + 1, 14].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 15]
                targetSheet.Cells[initalrow + 1, 15].Value2 = "";
                targetSheet.Cells[initalrow + 1, 15].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 15].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 15].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 15].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 16]
                targetSheet.Cells[initalrow + 1, 16].Value2 = "";
                targetSheet.Cells[initalrow + 1, 16].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 16].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 16].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 16].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 17]
                targetSheet.Cells[initalrow + 1, 17].Value2 = "";
                targetSheet.Cells[initalrow + 1, 17].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 17].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 17].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 17].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 18]
                targetSheet.Cells[initalrow + 1, 18].Value2 = "";
                targetSheet.Cells[initalrow + 1, 18].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 18].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 18].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 18].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 19]
                targetSheet.Cells[initalrow + 1, 19].Value2 = "";
                targetSheet.Cells[initalrow + 1, 19].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 19].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 19].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 19].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 20]
                targetSheet.Cells[initalrow + 1, 20].Value2 = "";
                targetSheet.Cells[initalrow + 1, 20].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 20].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 20].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 20].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 21]
                targetSheet.Cells[initalrow + 1, 21].Value2 = "";
                targetSheet.Cells[initalrow + 1, 21].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 21].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 21].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 21].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 22]
                targetSheet.Cells[initalrow + 1, 22].Value2 = "";
                targetSheet.Cells[initalrow + 1, 22].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 22].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 22].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 22].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 22].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 23]
                targetSheet.Cells[initalrow + 1, 23].Value2 = "";
                targetSheet.Cells[initalrow + 1, 23].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 23].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 23].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 23].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 23].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 24]
                targetSheet.Cells[initalrow + 1, 24].Value2 = "";
                targetSheet.Cells[initalrow + 1, 24].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 24].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 24].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 24].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 24].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            }
            targetSheet.Columns["K:W"].ColumnWidth = 10;
            targetSheet.Columns["A:A"].ColumnWidth = 15;
            targetSheet.Columns["X:X"].ColumnWidth = 20;
            Excel.Range range = targetSheet.Range[$"A{initrow}:Y{initrow+1}"];
            range.Merge();
            range. Value2 = "宴 会 营 业 收 入 输 入 表:";
            range.Font.Size = 15;
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Font.Bold = true;
            range = targetSheet.Range[$"A{initrow+2}:B{initrow + 2}"];
            range.Merge();
            range.Value2 = $"Month:{strmonth}";            
            int argb = int.Parse("00FFFF", System.Globalization.NumberStyles.HexNumber);
            Color color = Color.FromArgb(argb);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
            range.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            range.Font.Bold = true;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range = targetSheet.Range[$"C{initrow+2}:Y{initrow+2}"];
            range.Merge();
            targetSheet.Range[$"A{initalrow+2}"].Select();
            targetSheet.Application.ActiveWindow.FreezePanes = true;
        }
        private void ApplyZcHeaderToActiveSheet(Excel.Worksheet targetSheet, int initalrow = 4, string strmonth = "")
        {

            int cellcolor = 14150647;
            int initrow = initalrow - 3;
            // Cell [1, 1]
            {

                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 3], targetSheet.Cells[initalrow + 1, 3]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 5], targetSheet.Cells[initalrow + 1, 5]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 13], targetSheet.Cells[initalrow + 1, 13]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 13], targetSheet.Cells[initalrow + 1, 13]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 13], targetSheet.Cells[initalrow + 1, 13]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 15], targetSheet.Cells[initalrow + 1, 15]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 17], targetSheet.Cells[initalrow + 1, 17]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 2], targetSheet.Cells[initalrow + 1, 2]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 14], targetSheet.Cells[initalrow + 1, 14]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 14], targetSheet.Cells[initalrow + 1, 14]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 14], targetSheet.Cells[initalrow + 1, 14]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 20], targetSheet.Cells[initalrow + 1, 20]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 7], targetSheet.Cells[initalrow + 1, 7]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 9], targetSheet.Cells[initalrow + 1, 9]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 1], targetSheet.Cells[initalrow + 1, 1]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 19], targetSheet.Cells[initalrow + 1, 19]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 21], targetSheet.Cells[initalrow + 1, 21]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow, 12]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow, 12]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 10], targetSheet.Cells[initalrow, 12]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 4], targetSheet.Cells[initalrow + 1, 4]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 6], targetSheet.Cells[initalrow + 1, 6]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 8], targetSheet.Cells[initalrow + 1, 8]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 16], targetSheet.Cells[initalrow + 1, 16]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Merge();
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Interior.Color = 14150647;
                targetSheet.Range[targetSheet.Cells[initalrow, 18], targetSheet.Cells[initalrow + 1, 18]].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                // Cell [1, 1]
                targetSheet.Cells[initalrow, 1].Value2 = "日期";
                targetSheet.Cells[initalrow, 1].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 1].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 1].Font.Bold = true;
                targetSheet.Cells[initalrow, 1].Font.Italic = false;
                targetSheet.Cells[initalrow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 2]
                targetSheet.Cells[initalrow, 2].Value2 = "账单号";
                targetSheet.Cells[initalrow, 2].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 2].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 2].Font.Bold = true;
                targetSheet.Cells[initalrow, 2].Font.Italic = false;
                targetSheet.Cells[initalrow, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 2].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 3]
                targetSheet.Cells[initalrow, 3].Value2 = "部门";
                targetSheet.Cells[initalrow, 3].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 3].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 3].Font.Bold = true;
                targetSheet.Cells[initalrow, 3].Font.Italic = false;
                targetSheet.Cells[initalrow, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 3].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 4]
                targetSheet.Cells[initalrow, 4].Value2 = "预订人";
                targetSheet.Cells[initalrow, 4].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 4].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 4].Font.Bold = true;
                targetSheet.Cells[initalrow, 4].Font.Italic = false;
                targetSheet.Cells[initalrow, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 4].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 5]
                targetSheet.Cells[initalrow, 5].Value2 = "地点";
                targetSheet.Cells[initalrow, 5].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 5].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 5].Font.Bold = true;
                targetSheet.Cells[initalrow, 5].Font.Italic = false;
                targetSheet.Cells[initalrow, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 5].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 6]
                targetSheet.Cells[initalrow, 6].Value2 = "人数";
                targetSheet.Cells[initalrow, 6].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 6].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 6].Font.Bold = true;
                targetSheet.Cells[initalrow, 6].Font.Italic = false;
                targetSheet.Cells[initalrow, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 6].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 7]
                targetSheet.Cells[initalrow, 7].Value2 = "桌数";
                targetSheet.Cells[initalrow, 7].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 7].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 7].Font.Bold = true;
                targetSheet.Cells[initalrow, 7].Font.Italic = false;
                targetSheet.Cells[initalrow, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 7].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 8]
                targetSheet.Cells[initalrow, 8].Value2 = "餐标";
                targetSheet.Cells[initalrow, 8].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 8].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 8].Font.Bold = true;
                targetSheet.Cells[initalrow, 8].Font.Italic = false;
                targetSheet.Cells[initalrow, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 8].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 9]
                targetSheet.Cells[initalrow, 9].Value2 = "人均";
                targetSheet.Cells[initalrow, 9].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 9].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 9].Font.Bold = true;
                targetSheet.Cells[initalrow, 9].Font.Italic = false;
                targetSheet.Cells[initalrow, 9].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 10]
                targetSheet.Cells[initalrow, 10].Value2 = "餐段";
                targetSheet.Cells[initalrow, 10].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 10].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 10].Font.Bold = true;
                targetSheet.Cells[initalrow, 10].Font.Italic = false;
                targetSheet.Cells[initalrow, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 10].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 11]
                targetSheet.Cells[initalrow, 11].Value2 = "";
                targetSheet.Cells[initalrow, 11].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 11].Font.Size = 12.0;
                targetSheet.Cells[initalrow, 11].Font.Bold = false;
                targetSheet.Cells[initalrow, 11].Font.Italic = false;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 12]
                targetSheet.Cells[initalrow, 12].Value2 = "";
                targetSheet.Cells[initalrow, 12].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 12].Font.Size = 12.0;
                targetSheet.Cells[initalrow, 12].Font.Bold = false;
                targetSheet.Cells[initalrow, 12].Font.Italic = false;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 13]
                targetSheet.Cells[initalrow, 13].Value2 = "红酒";
                targetSheet.Cells[initalrow, 13].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 13].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 13].Font.Bold = true;
                targetSheet.Cells[initalrow, 13].Font.Italic = false;
                targetSheet.Cells[initalrow, 13].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 13].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 14]
                targetSheet.Cells[initalrow, 14].Value2 = "烈酒";
                targetSheet.Cells[initalrow, 14].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 14].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 14].Font.Bold = true;
                targetSheet.Cells[initalrow, 14].Font.Italic = false;
                targetSheet.Cells[initalrow, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 14].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 15]
                targetSheet.Cells[initalrow, 15].Value2 = "啤酒";
                targetSheet.Cells[initalrow, 15].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 15].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 15].Font.Bold = true;
                targetSheet.Cells[initalrow, 15].Font.Italic = false;
                targetSheet.Cells[initalrow, 15].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 15].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 16]
                targetSheet.Cells[initalrow, 16].Value2 = "软饮";
                targetSheet.Cells[initalrow, 16].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 16].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 16].Font.Bold = true;
                targetSheet.Cells[initalrow, 16].Font.Italic = false;
                targetSheet.Cells[initalrow, 16].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 16].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 17]
                targetSheet.Cells[initalrow, 17].Value2 = "杂项";
                targetSheet.Cells[initalrow, 17].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 17].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 17].Font.Bold = true;
                targetSheet.Cells[initalrow, 17].Font.Italic = false;
                targetSheet.Cells[initalrow, 17].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 17].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 18]
                targetSheet.Cells[initalrow, 18].Value2 = "合计";
                targetSheet.Cells[initalrow, 18].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 18].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 18].Font.Bold = true;
                targetSheet.Cells[initalrow, 18].Font.Italic = false;
                targetSheet.Cells[initalrow, 18].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 18].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 19]
                targetSheet.Cells[initalrow, 19].Value2 = "净价";
                targetSheet.Cells[initalrow, 19].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 19].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 19].Font.Bold = true;
                targetSheet.Cells[initalrow, 19].Font.Italic = false;
                targetSheet.Cells[initalrow, 19].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 19].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 20]
                targetSheet.Cells[initalrow, 20].Value2 = "结账方式";
                targetSheet.Cells[initalrow, 20].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 20].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 20].Font.Bold = true;
                targetSheet.Cells[initalrow, 20].Font.Italic = false;
                targetSheet.Cells[initalrow, 20].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 20].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [1, 21]
                targetSheet.Cells[initalrow, 21].Value2 = "checkTotal";
                targetSheet.Cells[initalrow, 21].Font.Name = "宋体";
                targetSheet.Cells[initalrow, 21].Font.Size = 9.0;
                targetSheet.Cells[initalrow, 21].Font.Bold = true;
                targetSheet.Cells[initalrow, 21].Font.Italic = false;
                targetSheet.Cells[initalrow, 21].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow, 21].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 1]
                targetSheet.Cells[initalrow + 1, 1].Value2 = "";
                targetSheet.Cells[initalrow + 1, 1].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 1].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 1].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 1].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 2]
                targetSheet.Cells[initalrow + 1, 2].Value2 = "";
                targetSheet.Cells[initalrow + 1, 2].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 2].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 2].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 2].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 3]
                targetSheet.Cells[initalrow + 1, 3].Value2 = "";
                targetSheet.Cells[initalrow + 1, 3].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 3].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 3].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 3].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 4]
                targetSheet.Cells[initalrow + 1, 4].Value2 = "";
                targetSheet.Cells[initalrow + 1, 4].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 4].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 4].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 4].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 4].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 5]
                targetSheet.Cells[initalrow + 1, 5].Value2 = "";
                targetSheet.Cells[initalrow + 1, 5].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 5].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 5].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 5].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 5].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 6]
                targetSheet.Cells[initalrow + 1, 6].Value2 = "";
                targetSheet.Cells[initalrow + 1, 6].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 6].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 6].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 6].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 7]
                targetSheet.Cells[initalrow + 1, 7].Value2 = "";
                targetSheet.Cells[initalrow + 1, 7].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 7].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 7].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 7].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 7].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 8]
                targetSheet.Cells[initalrow + 1, 8].Value2 = "";
                targetSheet.Cells[initalrow + 1, 8].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 8].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 8].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 8].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 8].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 9]
                targetSheet.Cells[initalrow + 1, 9].Value2 = "";
                targetSheet.Cells[initalrow + 1, 9].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 9].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 9].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 9].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 9].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 10]
                targetSheet.Cells[initalrow + 1, 10].Value2 = "早餐";
                targetSheet.Cells[initalrow + 1, 10].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 10].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 10].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 10].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 10].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 10].Interior.Color = 14150647;
                targetSheet.Cells[initalrow + 1, 10].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 10].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 11]
                targetSheet.Cells[initalrow + 1, 11].Value2 = "午餐";
                targetSheet.Cells[initalrow + 1, 11].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 11].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 11].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 11].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 11].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 11].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 11].Interior.Color = 14150647;
                targetSheet.Cells[initalrow + 1, 11].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 11].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 12]
                targetSheet.Cells[initalrow + 1, 12].Value2 = "晚餐";
                targetSheet.Cells[initalrow + 1, 12].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 12].Font.Size = 9.0;
                targetSheet.Cells[initalrow + 1, 12].Font.Bold = true;
                targetSheet.Cells[initalrow + 1, 12].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                targetSheet.Cells[initalrow + 1, 12].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                targetSheet.Cells[initalrow + 1, 12].Interior.Color = 14150647;
                targetSheet.Cells[initalrow + 1, 12].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 12].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 13]
                targetSheet.Cells[initalrow + 1, 13].Value2 = "红酒";
                targetSheet.Cells[initalrow + 1, 13].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 13].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 13].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 13].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 13].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 14]
                targetSheet.Cells[initalrow + 1, 14].Value2 = "烈酒";
                targetSheet.Cells[initalrow + 1, 14].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 14].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 14].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 14].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 14].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 15]
                targetSheet.Cells[initalrow + 1, 15].Value2 = "啤酒";
                targetSheet.Cells[initalrow + 1, 15].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 15].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 15].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 15].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 15].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 16]
                targetSheet.Cells[initalrow + 1, 16].Value2 = "软饮";
                targetSheet.Cells[initalrow + 1, 16].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 16].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 16].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 16].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 16].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 17]
                targetSheet.Cells[initalrow + 1, 17].Value2 = "杂项";
                targetSheet.Cells[initalrow + 1, 17].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 17].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 17].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 17].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 17].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 18]
                targetSheet.Cells[initalrow + 1, 18].Value2 = "合计";
                targetSheet.Cells[initalrow + 1, 18].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 18].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 18].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 18].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 18].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 19]
                targetSheet.Cells[initalrow + 1, 19].Value2 = "净价";
                targetSheet.Cells[initalrow + 1, 19].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 19].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 19].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 19].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 19].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 20]
                targetSheet.Cells[initalrow + 1, 20].Value2 = "结账方式";
                targetSheet.Cells[initalrow + 1, 20].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 20].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 20].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 20].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 20].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                // Cell [2, 21]
                targetSheet.Cells[initalrow + 1, 21].Value2 = "checkTotal";
                targetSheet.Cells[initalrow + 1, 21].Font.Name = "宋体";
                targetSheet.Cells[initalrow + 1, 21].Font.Size = 12.0;
                targetSheet.Cells[initalrow + 1, 21].Font.Bold = false;
                targetSheet.Cells[initalrow + 1, 21].Font.Italic = false;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                targetSheet.Cells[initalrow + 1, 21].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;


            }
            targetSheet.Columns["J:S"].ColumnWidth = 10;
            targetSheet.Columns["A:A"].ColumnWidth = 15;
            targetSheet.Columns["T:T"].ColumnWidth = 20;
            Excel.Range range = targetSheet.Range[$"A{initrow}:U{initrow + 1}"];
            range.Merge();
            range.Value2 = "中 餐 营 业 收 入 输 入 表:";
            range.Font.Size = 15;
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Font.Bold = true;
            range = targetSheet.Range[$"A{initrow + 2}:B{initrow + 2}"];
            range.Merge();
            range.Value2 = $"Month:{strmonth}";
            int argb = int.Parse("00FFFF", System.Globalization.NumberStyles.HexNumber);
            Color color = Color.FromArgb(argb);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
            range.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            range.Font.Bold = true;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range = targetSheet.Range[$"C{initrow + 2}:U{initrow + 2}"];
            range.Merge();
            targetSheet.Range[$"A{initalrow + 2}"].Select();
            targetSheet.Application.ActiveWindow.FreezePanes = true;
        }

        private void ApplySheetName(Excel.Worksheet targetSheet,string name)
        {
            if(targetSheet.Name!=  name)
                targetSheet.Name = name;
        }

 

        private static string GetColumnLetter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        private void banquetbtn_Click(object sender, RibbonControlEventArgs e)
        {
            string ls_date = banquetdate.Text.Trim();
            int liyear, limonth;
            string[] parts = ls_date.Split('-');
            liyear = int.Parse(parts[0]);
            limonth = int.Parse(parts[1]);
            if (!(limonth >= 1 && limonth <= 12&& liyear>=2024 && liyear <= 2034))
            {
                MessageBox.Show($"无{ls_date}无效！");
                return;
            }
            string[] lsip = getRegStr().Split(':');
            if (!lsip[1].Equals("3309"))
            {
                connectionString = $"Server={lsip[0]};Port={lsip[1]};Database=hotel;User ID=root;Password=arzfUh??p3<L;";
            }
            
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                List<BanquetData> dataList = ReceiveData.FetchBanquetDataByDate(conn,  ls_date);
                if (dataList.Count <= 0)
                {
                    MessageBox.Show($"无{ls_date}日数据！");
                    return;
                }
                long vaildmaxrows=1;
                Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Excel.Worksheet curworksheet=null;
                if (activeWorkbook.Sheets.Count > 0)
                {
                    foreach (Worksheet sheet in activeWorkbook.Sheets)
                    {
                        if (sheet.Name == "宴会"+ls_date)
                        {
                            curworksheet = sheet;
                            curworksheet.Activate();
                            break;
                        }
                    }
                }
                if(curworksheet == null)
                {
                    curworksheet = activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[activeWorkbook.Sheets.Count]);
                }
                Excel.Range range1 = curworksheet.Range["A4:A5"];
                Excel.Range range2 = curworksheet.Range["B4:B5"];
                if(!(range1.MergeCells && range2.MergeCells && range1.Cells[1, 1].Value2.ToString()??"" == "Date"
                    && range2.Cells[1, 1].Value2.ToString()??"" == "Check No")) {
                    ApplyBanquetHeaderToActiveSheet(curworksheet, 4, ls_date);
                    ApplySheetName(curworksheet, "宴会" + ls_date);
                }
                vaildmaxrows = curworksheet.UsedRange.Rows.Count+1;
                long initrow = vaildmaxrows;
                for (int i = 0; i < dataList.Count; i++)
                {
                    curworksheet.Cells[i + vaildmaxrows, "A"] = dataList[i].openDateTime;
                    curworksheet.Cells[i+ vaildmaxrows, "B"] = dataList[i].checkNum ;
                    curworksheet.Cells[i + vaildmaxrows, "G"] = dataList[i].numGuests.ToString()??"";
                    curworksheet.Cells[i + vaildmaxrows, "H"] = dataList[i].tablescount.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "I"] = dataList[i].tablesper.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "K"] = dataList[i].Breakfast.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "L"] = dataList[i].Lunch.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "M"] = dataList[i].Dinner.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "N"] = dataList[i].chaxie.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "O"] = dataList[i].bever.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "P"] = dataList[i].misce.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "Q"] = dataList[i].beer.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "R"] = dataList[i].equiment.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "S"] = dataList[i].room.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "V"] = $"= round(SUBTOTAL(9, K{i + vaildmaxrows}: U{i + vaildmaxrows}),2)";
                    //worksheet.Cells[dataList.Count + max_row, "D"] = $"= SUBTOTAL(9, D{inital_rows}: D{dataList.Count + max_row - 1})";
                    curworksheet.Cells[i + vaildmaxrows, "W"] = $"=round(V{i + vaildmaxrows}/1.06,2)";
                    curworksheet.Cells[i + vaildmaxrows, "X"] = dataList[i].paymethod;
                    curworksheet.Cells[i + vaildmaxrows, "Y"] = dataList[i].checkTotal??0;

                    Excel.Range rangeV = curworksheet.Cells[i + vaildmaxrows, "V"];
                    rangeV.FormatConditions.Delete();
                    Excel.Range rangeY = curworksheet.Cells[i + vaildmaxrows, "Y"];
                    // 创建条件格式规则
                    Excel.FormatCondition condition = curworksheet.Range[$"V{i + vaildmaxrows}"].FormatConditions.Add(
                        Excel.XlFormatConditionType.xlCellValue,
                        Excel.XlFormatConditionOperator.xlEqual,
                        rangeY.Value
                    );
                    condition.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E2EFDA"));
                }

                for (int col = 7; col <= 23; col++)
                {
                    string colLetter = GetColumnLetter(col);
                    if (!(colLetter == "I" || colLetter == "J"))
                    {
                        string formula = $"=sum( {colLetter}{initrow}:{colLetter}{dataList.Count + initrow - 1})";
                        curworksheet.Cells[dataList.Count + initrow, col] = formula;
                    }
                }

                range1 = curworksheet.Range[$"A{dataList.Count + vaildmaxrows}:F{dataList.Count + vaildmaxrows}"];
                range1.Merge();
                range1.Value2 = "汇总";
                range1.Font.Bold = true;
                range1.Font.Size = 14;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                curworksheet.Range[$"A{dataList.Count + initrow}:Y{dataList.Count + initrow}"].Font.Bold = true;
                curworksheet.Range[$"A{dataList.Count + initrow}:Y{dataList.Count + initrow}"].
                    Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E7E6E6"));
                curworksheet.Range[$"A{dataList.Count + vaildmaxrows}:Y{dataList.Count + vaildmaxrows}"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                curworksheet.Cells[vaildmaxrows, 1].Select();
                curworksheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                curworksheet.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                //curworksheet.Columns["A:Y"].AutoFit();
            }
            MessageBox.Show("数据生成完毕！");
            return;
        }

        private void chinesefoodbtn_Click(object sender, RibbonControlEventArgs e)
        {
            string ls_date = chinesefoodeditBox.Text.Trim();
            int liyear, limonth;
            string[] parts = ls_date.Split('-');
            liyear = int.Parse(parts[0]);
            limonth = int.Parse(parts[1]);
            if (!(limonth >= 1 && limonth <= 12 && liyear >= 2024 && liyear <= 2034))
            {
                MessageBox.Show($"无{ls_date}无效！");
                return;
            }
            string[] lsip = getRegStr().Split(':');
            if (!lsip[1].Equals("3309"))
            {
                connectionString = $"Server={lsip[0]};Port={lsip[1]};Database=hotel;User ID=root;Password=arzfUh??p3<L;";
            }

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                List<ChineseFoodData> dataList = ReceiveData.FetchChineseFoodDataByDate(conn, ls_date);
                if (dataList.Count <= 0)
                {
                    MessageBox.Show($"无{ls_date}日数据！");
                    return;
                }
                long vaildmaxrows = 1;
                Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Excel.Worksheet curworksheet = null;
                if (activeWorkbook.Sheets.Count > 0)
                {
                    foreach (Worksheet sheet in activeWorkbook.Sheets)
                    {
                        if (sheet.Name == "中餐" + ls_date)
                        {
                            curworksheet = sheet;
                            curworksheet.Activate();
                            break;
                        }
                    }
                }
                if (curworksheet == null)
                {
                    curworksheet = activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[activeWorkbook.Sheets.Count]);
                }
                Excel.Range range1 = curworksheet.Range["A4:A5"];
                Excel.Range range2 = curworksheet.Range["B4:B5"];
                if (!(range1.MergeCells && range2.MergeCells && range1.Cells[1, 1].Value2.ToString() == "日期"
                    && range2.Cells[1, 1].Value2.ToString() == "账单号"))
                {
                    ApplyZcHeaderToActiveSheet(curworksheet, 4, ls_date);
                    ApplySheetName(curworksheet, "中餐" + ls_date);
                }
                vaildmaxrows = curworksheet.UsedRange.Rows.Count + 1;
                long initrow = vaildmaxrows;
                for (int i = 0; i < dataList.Count; i++)
                {
                    curworksheet.Cells[i + vaildmaxrows, "A"] = dataList[i].openDateTime;
                    curworksheet.Cells[i + vaildmaxrows, "B"] = dataList[i].checkNum;
                    curworksheet.Cells[i + vaildmaxrows, "F"] = dataList[i].numGuests.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "G"] = dataList[i].tablescount.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "H"] = dataList[i].tablesper.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "I"] = $"=round(H{i + vaildmaxrows}/IF(AND(ISNUMBER(F{i + vaildmaxrows}) ,F{i + vaildmaxrows} > 0),F{i + vaildmaxrows},1),2)"; 
                    curworksheet.Cells[i + vaildmaxrows, "J"] = dataList[i].Breakfast.ToString() ?? ""; 
                    curworksheet.Cells[i + vaildmaxrows, "K"] = dataList[i].Lunch.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "L"] = dataList[i].Dinner.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "M"] = dataList[i].wine.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "N"] = dataList[i].liquor.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "O"] = dataList[i].beer.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "P"] = dataList[i].bever.ToString() ?? "";
                    curworksheet.Cells[i + vaildmaxrows, "Q"] = dataList[i].misce.ToString() ?? "";

                    curworksheet.Cells[i + vaildmaxrows, "R"] = $"= round(SUBTOTAL(9, J{i + vaildmaxrows}: Q{i + vaildmaxrows}),2)";
                    curworksheet.Cells[i + vaildmaxrows, "S"] = $"=round(R{i + vaildmaxrows}/1.06,2)";
                    curworksheet.Cells[i + vaildmaxrows, "T"] = dataList[i].paymethod;
                    curworksheet.Cells[i + vaildmaxrows, "U"] = dataList[i].checkTotal ?? 0;

                    Excel.Range rangeV = curworksheet.Cells[i + vaildmaxrows, "R"];
                    rangeV.FormatConditions.Delete();
                    Excel.Range rangeY = curworksheet.Cells[i + vaildmaxrows, "U"];
                    Excel.FormatCondition condition = curworksheet.Range[$"R{i + vaildmaxrows}"].FormatConditions.Add(
                        Excel.XlFormatConditionType.xlCellValue,
                        Excel.XlFormatConditionOperator.xlEqual,
                        rangeY.Value
                    );
                    condition.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E2EFDA"));
                }

                for (int col = 6; col <= 21; col++)// 汇总列
                {
                    string colLetter = GetColumnLetter(col);
                    if (colLetter != "T") { 
                        string formula = $"=sum( {colLetter}{initrow}:{colLetter}{dataList.Count + initrow - 1})";
                        curworksheet.Cells[dataList.Count + initrow, col] = formula;
                    }

                }

                range1 = curworksheet.Range[$"A{dataList.Count + vaildmaxrows}:E{dataList.Count + vaildmaxrows}"];
                range1.Merge();
                range1.Value2 = "汇总";
                range1.Font.Bold = true;
                range1.Font.Size = 14;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                curworksheet.Range[$"A{dataList.Count + initrow}:U{dataList.Count + initrow}"].Font.Bold = true;
                curworksheet.Range[$"A{dataList.Count + initrow}:U{dataList.Count + initrow}"].
                    Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E7E6E6"));
                curworksheet.Range[$"A{dataList.Count + vaildmaxrows}:U{dataList.Count + vaildmaxrows}"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                curworksheet.Cells[vaildmaxrows, 1].Select();
                curworksheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                curworksheet.UsedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                //curworksheet.Columns["A:Y"].AutoFit();
            }
            MessageBox.Show("数据生成完毕！");
            return;
        }
    }
}
