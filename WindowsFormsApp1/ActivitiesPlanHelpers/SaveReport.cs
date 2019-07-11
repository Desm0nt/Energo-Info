using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using unvell.ReoGrid;

namespace WindowsFormsApp1.ActivitiesPlanHelpers
{
    public partial class SaveReport : Form
    {
        private Worksheet worksheet;
        int _type;

        public SaveReport(int type, int year, string org, DataTable dt1, DataTable dt2)
        {
            InitializeComponent();
            dt1 = Round(dt1);
            dt2 = Round(dt2);
            string s = "";
            _type = type;
            int w = 0;
            int start = 0;
            int total1 = 0, total2 = 0, total3 = 0, total = 0;
            dt1.Rows.RemoveAt(dt1.Rows.Count - 1);
            if (dt1.Columns["ID"] != null)
                dt1.Columns.Remove("ID");
            dt2.Rows.RemoveAt(dt2.Rows.Count - 1);
            dt2.Rows.RemoveAt(dt2.Rows.Count - 1);
            if (dt2.Columns["ID"] != null)
                dt2.Columns.Remove("ID");

            if (type == 1)
            {
                if (dt1.Columns["Дата внедрения"] != null)
                    dt1.Columns.Remove("Дата внедрения");
                if (dt2.Columns["Дата внедрения"] != null)
                    dt2.Columns.Remove("Дата внедрения");
                reoGridControl1.Load(Directory.GetCurrentDirectory() + "\\ActivitiesPlanHelpers\\ReportForms\\Report1.xlsx");

                

                //reoGridControl1.Load("..\\..\\Reports\\1.xlsx");
                start = 7;
                w = 5;
            }
            if (type == 2)
            {
                reoGridControl1.Load(Directory.GetCurrentDirectory() + "\\ActivitiesPlanHelpers\\ReportForms\\Report2_1.xlsx");
                //reoGridControl1.Load("..\\..\\Reports\\1.xlsx");
                start = 8;
                w = 4;
            }

            worksheet = reoGridControl1.CurrentWorksheet;
            worksheet.SelectionStyle = WorksheetSelectionStyle.None;
            worksheet.SetSettings(WorksheetSettings.Behavior_MouseWheelToZoom, false);

            //todo
            if (type == 1)
            {
                worksheet["A1"] = "Перечень мероприятий по энергосбережению на " + year.ToString() + " год по " + org;
                s = "FGHIJKLMNOPQRSTUVW";

                worksheet.SetRowsHeight(4, 1, 85);
                worksheet.SetColumnsWidth(0, 2, 50);
                worksheet.SetColumnsWidth(2, 1, 250);
                worksheet.SetColumnsWidth(3, 25, 60);
                worksheet["S3"] = year;
                //FillZero(7, 5, 23);
                //FillZero(9, 5, 23);
                //FillZero(11, 5, 23);
                //FillZero(12, 5, 23);
            }
            if (type == 2)
            {
                worksheet["A2"] = "нетрадиционных и возобновляемых энергоресурсов на " + year.ToString() + " год (поквартальная разбивка) по " + org;
                s = "EFGHIJKLMNOPQRSTUV";
                worksheet.SetRangeStyles("A2:W2", new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName,
                    FontName = "Times New Roman",
                    FontSize = 10,
                });
                worksheet.SetRangeStyles("A2:W7", new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.TextWrap,
                    TextWrapMode = TextWrapMode.WordBreak,

                });
                worksheet.SetRowsHeight(6, 1, 85);
                //worksheet.SetColumnsWidth(0, 2, 50);
                //worksheet.SetColumnsWidth(2, 1, 250);
                //worksheet.SetColumnsWidth(3, 25, 60);
                worksheet["R4"] = year;
            }

            
            Fill(s, dt1, ref start, ref total, false, w);
            total1 = total;
            Fill(s, dt1, ref start, ref total, true, w);
            total2 = total;
            Fill(s, dt2, ref start, ref total, w);
            total3 = total;
            for (int i = 0; i < s.Length; i++)
            {
                string ss = "" + s[i] + (total1) + "+" + s[i] + (total2) + "+" + s[i] + (total3);
                worksheet.Cells[total, i + w].Formula = ss;
            }
            worksheet.SetRangeStyles("D" + (total+1) + ":W" + (total+1), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName,
                FontName = "Arial",
                FontSize = 10,
            });


            сохранитьКакToolStripMenuItem.PerformClick();
            this.Close();
            

        }

        private void FillZero(int n, int start, int end)
        {
            for (int j = start; j < end; j++)
            {
                worksheet[n, j] = 0.0f;
            }
        }
        private void Fill(string s, DataTable dt, ref int start, ref int total, bool flag, int w)
        {
            int k = 0;




            for (int j = 0; j < dt.Rows.Count; j++)
            {
                if (Convert.ToBoolean(dt.Rows[j]["Доп. прогр."]) == flag)
                {
                    worksheet.InsertRows(start + k, 1);
                    
                    for (int i = 0; i < dt.Columns.Count - 1; i++)
                    {
                        if (_type == 2 && i == 3)

                            worksheet[start + k, i] = dt.Rows[j][i].ToString();
                        else
                        {
                            float f;
                            if (float.TryParse(dt.Rows[j][i].ToString(), out f))
                                worksheet[start + k, i] = Math.Round(f, 1);
                            else
                                worksheet[start + k, i] = dt.Rows[j][i];
                        }
                        worksheet[start + k, i] = dt.Rows[j][i];
                    }
                    k++;
                }
            }
            total = start + k;

            worksheet.SetRangeStyles("A" + (start+1) + ":W" + (total), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName | PlainStyleFlag.TextWrap,
                FontName = "Arial",
                FontSize = 9,
                TextWrapMode = TextWrapMode.WordBreak,
            });

            //worksheet.SetRangeStyles("C" + (start + 1) + ":C" + (total), new WorksheetRangeStyle
            //{
            //    Flag = PlainStyleFlag.TextWrap,
            //    TextWrapMode = TextWrapMode.WordBreak,
            //});


            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < s.Length; i++)
                {
                    worksheet.Cells[total, i + w].Formula = "SUM(" + s[i] + (start + 1) + ":" + s[i] + (start + k) + ")";
                }
            }

            start = total + 2;

            worksheet.SetRangeStyles("D" + (start-1) + ":W" + (start-1), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName,
                FontName = "Arial",
                FontSize = 10,
            });

            total++;
        }
        private void Fill(string s, DataTable dt, ref int start, ref int total, int w)
        {
            int k = 0;
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                worksheet.InsertRows(start + k, 1);
                for (int i = 0; i < dt.Columns.Count - 1; i++)
                {
                    float f;
                    if (float.TryParse(dt.Rows[j][i].ToString(), out f))
                        worksheet[start + k, i] = Math.Round(f, 1);
                    else
                        worksheet[start + k, i] = dt.Rows[j][i].ToString();
                }
                k++;
            }
            total = start + k;

            worksheet.SetRangeStyles("A" + (start + 1) + ":W" + (total), new WorksheetRangeStyle
            {

                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName | PlainStyleFlag.TextWrap,
                FontName = "Arial",
                FontSize = 9,
                TextWrapMode = TextWrapMode.WordBreak,
            });

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < s.Length; i++)
                {
                    worksheet.Cells[total, i + w].Formula = "SUM(" + s[i] + (start + 1) + ":" + s[i] + (start + k) + ")";
                }
            }
            start = total + 2;

            worksheet.SetRangeStyles("D" + (start - 1) + ":W" + (start - 1), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName,
                FontName = "Arial",
                FontSize = 10,
            });

            total++;
        }
        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var workbook = reoGridControl1;
                //workbook.Worksheets[1] = worksheet;
                workbook.Worksheets[0].Name = "Приложение1";
                workbook.Save(Directory.GetCurrentDirectory() + "\\Reports" + "_" + DateTime.Today.ToString("yyyy") + "_" + DateTime.Today.ToString("MMMM") + ".xlsx", unvell.ReoGrid.IO.FileFormat.Excel2007);
                System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\Reports" + "_" + DateTime.Today.ToString("yyyy") + "_" + DateTime.Today.ToString("MMMM") + ".xlsx");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
        private DataTable Round(DataTable dt)
        {
            foreach (DataColumn dc in dt.Columns)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    float f;
                    if (float.TryParse(dr[dc].ToString(), out f))
                    {
                        dr[dc] =Math.Round(f,1);
                    }
                }
            }
            return dt;
        }
    }
}
