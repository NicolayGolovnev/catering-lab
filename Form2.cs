using OfficeOpenXml;
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

namespace Catering
{
    public partial class Form2 : Form
    {
        public Signatures signatures = new Signatures();
        public Dictionary<String, Int32> names = new Dictionary<string, int>
        {
            { "Свекла", 1 },
            { "Картофель", 2 },
            { "Подсолнечник", 3 },
            { "Лен", 4 },
            { "Хлопок", 5 },
            { "Молоко", 6 },
            { "Кожа", 7 },
            { "Шерсть", 8 },
            { "Животный жир", 9 }
        };

        public Dictionary<String, Int32> units = new Dictionary<string, int>
        {
            { "мл", 111 },
            { "л", 112 },
            { "г", 163 },
            { "кг", 166 },
            { "т", 168 }
        };
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) 
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel document (*.xlsx, *.xls)|*.xlsx;*.xls";
            dialog.Title = "Сохранение документа";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(@"..\..\xlsx\LAW_26677_attach_LAW_26677_5.xlsx")))
                {
                    string[] buf;
                    var worksheet = excelPackage.Workbook.Worksheets[0];

                    worksheet.Cells["AV6"].LoadFromText(textBox13.Text);
                    worksheet.Cells["AV7"].LoadFromText(textBox14.Text);
                    worksheet.Cells["AV10"].LoadFromText(textBox15.Text);
                    worksheet.Cells["AV11"].LoadFromText(textBox16.Text);
                    worksheet.Cells["A6"].LoadFromText(textBox2.Text);
                    worksheet.Cells["A9"].LoadFromText(textBox3.Text);

                    worksheet.Cells["AR14"].LoadFromText(signatures.stringBox8);
                    worksheet.Cells["AR16"].LoadFromText(signatures.stringBox10);
                    worksheet.Cells["AW16"].LoadFromText(signatures.stringBox9);
                    
                    try
                    {
                        buf = signatures.stringDate2.Split('.');
                    }
                    catch (Exception exp)
                    {
                        buf = new string[] { "0", "01", "00" };
                    }
                    worksheet.Cells["AS18"].LoadFromText(buf[0]);
                    if (buf[1] == "01")
                        buf[1] = "января";
                    else if (buf[1] == "02")
                        buf[1] = "февраля";
                    else if (buf[1] == "03")
                        buf[1] = "марта";
                    else if (buf[1] == "04")
                        buf[1] = "апреля";
                    else if (buf[1] == "05")
                        buf[1] = "мая";
                    else if (buf[1] == "06")
                        buf[1] = "июня";
                    else if (buf[1] == "07")
                        buf[1] = "июля";
                    else if (buf[1] == "08")
                        buf[1] = "августа";
                    else if (buf[1] == "09")
                        buf[1] = "сентября";
                    else if (buf[1] == "10")
                        buf[1] = "октября";
                    else if (buf[1] == "11")
                        buf[1] = "ноября";
                    else if (buf[1] == "12")
                        buf[1] = "декабря";
                    worksheet.Cells["AV18"].LoadFromText(buf[1]);
                    worksheet.Cells["BB18"].LoadFromText(buf[2].Substring(buf[2].Length - 2));

                    worksheet.Cells["AC18"].LoadFromText(textBox12.Text);
                    worksheet.Cells["AJ18"].LoadFromText(dateTimePicker6.Text);

                    worksheet.Cells["A20"].LoadFromText(textBox7.Text);
                    worksheet.Cells["E22"].LoadFromText(textBox5.Text);
                    worksheet.Cells["P22"].LoadFromText(textBox4.Text);
                    worksheet.Cells["F24"].LoadFromText(textBox6.Text);

                    int rowExl = 32;
                    //dataGrid
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1[0, i].Value == null)
                            worksheet.Cells[string.Format("A{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("A{0}", rowExl + i)].LoadFromText(dataGridView1[0, i].Value.ToString());
                        
                        if (dataGridView1[1, i].Value == null)
                            worksheet.Cells[string.Format("V{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("V{0}", rowExl + i)].LoadFromText(dataGridView1[1, i].Value.ToString());
                        
                        if (dataGridView1[2, i].Value == null)
                            worksheet.Cells[string.Format("Z{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("Z{0}", rowExl + i)].LoadFromText(dataGridView1[2, i].Value.ToString());

                        if (dataGridView1[3, i].Value == null)
                            worksheet.Cells[string.Format("AD{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("AD{0}", rowExl + i)].LoadFromText(dataGridView1[3, i].Value.ToString());

                        if (dataGridView1[4, i].Value == null)
                            worksheet.Cells[string.Format("AI{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("AI{0}", rowExl + i)].LoadFromText(dataGridView1[4, i].Value.ToString());

                        if (dataGridView1[5, i].Value == null)
                            worksheet.Cells[string.Format("AN{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("AN{0}", rowExl + i)].LoadFromText(dataGridView1[5, i].Value.ToString());

                        if (dataGridView1[6, i].Value == null)
                            worksheet.Cells[string.Format("AV{0}", rowExl + i)].LoadFromText("");
                        else
                            worksheet.Cells[string.Format("AV{0}", rowExl + i)].LoadFromText(dataGridView1[6, i].Value.ToString());
                    }
                    worksheet.Cells["AV59"].LoadFromText(textBox11.Text);

                    // Оборотная сторона
                    buf = textBox11.Text.Split(',');
                    worksheet.Cells["A66"].LoadFromText(RusNumber.Str(Int32.Parse(buf[0])));
                    if (buf.Length == 1)
                        worksheet.Cells["AW66"].LoadFromText("0");
                    else
                        worksheet.Cells["AW66"].LoadFromText(buf[1].ToString());
                    worksheet.Cells["I68"].LoadFromText(textBox19.Text);
                    worksheet.Cells["R68"].LoadFromText(textBox20.Text);
                    worksheet.Cells["A70"].LoadFromText(textBox21.Text);
                    try
                    {
                        buf = dateTimePicker3.Text.Split('.');
                    }
                    catch (Exception exp)
                    {
                        buf = new string[] { "0", "01", "00" };
                    }
                    worksheet.Cells["AO70"].LoadFromText(buf[0]);
                    if (buf[1] == "01")
                        buf[1] = "января";
                    else if (buf[1] == "02")
                        buf[1] = "февраля";
                    else if (buf[1] == "03")
                        buf[1] = "марта";
                    else if (buf[1] == "04")
                        buf[1] = "апреля";
                    else if (buf[1] == "05")
                        buf[1] = "мая";
                    else if (buf[1] == "06")
                        buf[1] = "июня";
                    else if (buf[1] == "07")
                        buf[1] = "июля";
                    else if (buf[1] == "08")
                        buf[1] = "августа";
                    else if (buf[1] == "09")
                        buf[1] = "сентября";
                    else if (buf[1] == "10")
                        buf[1] = "октября";
                    else if (buf[1] == "11")
                        buf[1] = "ноября";
                    else if (buf[1] == "12")
                        buf[1] = "декабря";
                    worksheet.Cells["AR70"].LoadFromText(buf[1]);
                    worksheet.Cells["BA70"].LoadFromText(buf[2].Substring(buf[2].Length - 2));
                    worksheet.Cells["A74"].LoadFromText(textBox22.Text);

                    worksheet.Cells["Z77"].LoadFromText(textBox24.Text);
                    worksheet.Cells["A79"].LoadFromText(textBox23.Text);
                    try
                    {
                        buf = dateTimePicker4.Text.Split('.');
                    }
                    catch (Exception exp)
                    {
                        buf = new string[] { "0", "01", "00" };
                    }
                    worksheet.Cells["AO79"].LoadFromText(buf[0]);
                    if (buf[1] == "01")
                        buf[1] = "января";
                    else if (buf[1] == "02")
                        buf[1] = "февраля";
                    else if (buf[1] == "03")
                        buf[1] = "марта";
                    else if (buf[1] == "04")
                        buf[1] = "апреля";
                    else if (buf[1] == "05")
                        buf[1] = "мая";
                    else if (buf[1] == "06")
                        buf[1] = "июня";
                    else if (buf[1] == "07")
                        buf[1] = "июля";
                    else if (buf[1] == "08")
                        buf[1] = "августа";
                    else if (buf[1] == "09")
                        buf[1] = "сентября";
                    else if (buf[1] == "10")
                        buf[1] = "октября";
                    else if (buf[1] == "11")
                        buf[1] = "ноября";
                    else if (buf[1] == "12")
                        buf[1] = "декабря";
                    worksheet.Cells["AR79"].LoadFromText(buf[1]);
                    worksheet.Cells["BA79"].LoadFromText(buf[2].Substring(buf[2].Length - 2));
                    worksheet.Cells["E81"].LoadFromText(textBox25.Text);
                    worksheet.Cells["AR81"].LoadFromText(textBox26.Text);
                    
                    worksheet.Cells["Y84"].LoadFromText(textBox27.Text);
                    worksheet.Cells["A88"].LoadFromText(textBox28.Text);
                    try
                    {
                        buf = dateTimePicker5.Text.Split('.');
                    }
                    catch (Exception exp)
                    {
                        buf = new string[] { "0", "01", "00" };
                    }
                    worksheet.Cells["AO88"].LoadFromText(buf[0]);
                    if (buf[1] == "01")
                        buf[1] = "января";
                    else if (buf[1] == "02")
                        buf[1] = "февраля";
                    else if (buf[1] == "03")
                        buf[1] = "марта";
                    else if (buf[1] == "04")
                        buf[1] = "апреля";
                    else if (buf[1] == "05")
                        buf[1] = "мая";
                    else if (buf[1] == "06")
                        buf[1] = "июня";
                    else if (buf[1] == "07")
                        buf[1] = "июля";
                    else if (buf[1] == "08")
                        buf[1] = "августа";
                    else if (buf[1] == "09")
                        buf[1] = "сентября";
                    else if (buf[1] == "10")
                        buf[1] = "октября";
                    else if (buf[1] == "11")
                        buf[1] = "ноября";
                    else if (buf[1] == "12")
                        buf[1] = "декабря";
                    worksheet.Cells["AR88"].LoadFromText(buf[1]);
                    worksheet.Cells["BA88"].LoadFromText(buf[2].Substring(buf[2].Length - 2));
                    worksheet.Cells["E90"].LoadFromText(textBox29.Text);

                    worksheet.Cells["A96"].LoadFromText(RusNumber.Str(Convert.ToInt32(textBox36.Text)));
                    worksheet.Cells["A100"].LoadFromText(RusNumber.Str(Convert.ToInt32(textBox34.Text)));
                    worksheet.Cells["AW100"].LoadFromText(textBox33.Text);
                    
                    worksheet.Cells["J102"].LoadFromText(signatures.stringBox31);
                    worksheet.Cells["S102"].LoadFromText(signatures.stringBox30);
                    worksheet.Cells["K106"].LoadFromText(signatures.stringBox32);

                    excelPackage.SaveAs(dialog.FileName);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            signatures.ShowDialog();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("{0}-{1}", e.RowIndex, e.ColumnIndex);
            int row = e.RowIndex, col = e.ColumnIndex;
            if (row == -1)
                return;
            //Dictionaries
            //if (dataGridView1.Rows[row].Cells[0].Value != null)
            if (dataGridView1[0, row].Value != null)
                if (names.ContainsKey(dataGridView1[0, row].Value.ToString()))
                    dataGridView1[1, row].Value = names[dataGridView1[0, row].Value.ToString()];
            if (dataGridView1[2, row].Value != null)
                if (units.ContainsKey(dataGridView1[2, row].Value.ToString()))
                    dataGridView1[3, row].Value = units[dataGridView1[2, row].Value.ToString()];

            if (dataGridView1[4, row].Value != null && dataGridView1[5, row].Value != null)
                dataGridView1[6, row].Value = Convert.ToInt32(dataGridView1[4, row].Value) * Convert.ToDouble(dataGridView1[5, row].Value);

            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                sum += Convert.ToDouble(dataGridView1[6, i].Value);
            textBox11.Text = sum.ToString();
            string[] buf = textBox11.Text.Split(',');
            textBox17.Text = RusNumber.Str(Convert.ToInt32(buf[0]));
        }
    }
}
