using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

//using System.Reflection;
//using Excel = Microsoft.Office.Interop.Excel;

using System.IO;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Util;
using NPOI.HSSF.Util;

namespace EAT
{
    public partial class Form1 : Form
    {

        //
        private bool isFromExcel;
        private bool isFromTexts;
        private bool isFromAny;
        private DataTable dt;

        public Form1()
        {
            InitializeComponent();

            //
            dt = new DataTable();
            this.dataGridView1.DataSource = dt;
            dt.Columns.Add(new DataColumn("文件名"));
            dt.Columns.Add(new DataColumn("内容"));
            this.dataGridView1.Columns[1].Width = 300;

            isFromExcel = false;
            isFromTexts = false;
            isFromAny = false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();

            open.Filter = "Excel 工作簿|*.xlsx|Excel 97-2003 工作簿|*.xls";

            if (open.ShowDialog() == DialogResult.OK)
            {
                string filePath = open.FileName;

                //解析excel
                //LoadExcel(filePath);

                //解析Excel NPIO
                LoadExcelNPIO(filePath);
            }

            isFromExcel = true;
            isFromTexts = false;
            isFromAny = false;
        }

        private void LoadExcelNPIO(string filePath)
        {
            FileStream fs = File.OpenRead(filePath);
            IWorkbook workbook;

            // 2007版本  
            if (filePath.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook(fs);
            // 2003版本  
            else if (filePath.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook(fs);
            else
                return;

            ISheet sheet = workbook.GetSheetAt(0);

            dt.Clear();
            int iRowCount = sheet.LastRowNum;
            for (int i = 0; i < iRowCount; ++i)
            {
                IRow row = sheet.GetRow(i);
                DataRow dr = dt.NewRow();
                dr[0] = row.GetCell(0);
                dr[1] = row.GetCell(1);

                dt.Rows.Add(dr);
            }

            workbook.Close();
            fs.Close();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("版权说明: \n\n本软件由中山大学肖鹏博士团队开发\n设计策划：肖鹏，联系方式：cifangyue@163.com\n技术负责：王海明 肖鹏");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (isFromExcel)//excel->text
            {
                FolderBrowserDialog folder = new FolderBrowserDialog();
                if (folder.ShowDialog() == DialogResult.OK)
                {
                    string dirPath = folder.SelectedPath;

                    SaveFiles(dirPath);
                }
            }
            if (isFromTexts)//text->excel
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "Excel 工作簿|*.xls";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    string filePath = save.FileName;
                    SaveExcel(filePath);
                }
            }
            if (isFromAny)//anyfile->excel
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "Excel 工作簿|*.xls";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    string filePath = save.FileName;
                    SaveExcelAny(filePath);
                }
            }
        }

        private void SaveExcelAny(string filePath)
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet0");

            int rowCount = dt.Rows.Count;
            for (int i = 0; i < rowCount; ++i)
            {
                IRow row = sheet.CreateRow(i);
                ICell cellName = row.CreateCell(0);
                cellName.SetCellValue(dt.Rows[i][0].ToString());
                

                ICell cellHyperlink = row.CreateCell(1);
                cellHyperlink.SetCellValue("<链接>");

                //超链接
                HSSFHyperlink Hyperlink = new HSSFHyperlink(HyperlinkType.Url);
                Hyperlink.Address =  dt.Rows[i][0].ToString();
                cellHyperlink.Hyperlink = Hyperlink;
                //颜色+下划线
                IFont font = workbook.CreateFont();//创建字体样式  
                font.Color = HSSFColor.Blue.Index;//设置字体颜色  
                font.Underline = FontUnderlineType.Single;//下划线
                ICellStyle style = workbook.CreateCellStyle();//创建单元格样式  
                style.SetFont(font);//设置单元格样式中的字体样式  
                cellHyperlink.CellStyle = style;//为单元格设置显示样式  
            }

            using (FileStream fs = File.Create(filePath))
            {
                workbook.Write(fs);
                workbook.Close();
                fs.Close();
            }

            MessageBox.Show("转换完成");
        }

        private void SaveExcel(string filePath)
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet0");

            int rowCount = dt.Rows.Count;
            for (int i = 0; i < rowCount; ++i)
            {
                IRow row = sheet.CreateRow(i);
                ICell cellName = row.CreateCell(0);
                cellName.SetCellValue(dt.Rows[i][0].ToString());
                ICell cellContent = row.CreateCell(1);
                cellContent.SetCellValue(dt.Rows[i][1].ToString());

                ICell cellHyperlink = row.CreateCell(2);
                cellHyperlink.SetCellValue("<链接>");

                //超链接
                HSSFHyperlink Hyperlink = new HSSFHyperlink(HyperlinkType.File);
                Hyperlink.Address = "./" + dt.Rows[i][0].ToString() + ".txt";
                cellHyperlink.Hyperlink = Hyperlink;

                //颜色+下划线
                IFont font = workbook.CreateFont();//创建字体样式  
                font.Color = HSSFColor.Blue.Index;//设置字体颜色  
                font.Underline = FontUnderlineType.Single;//下划线
                ICellStyle style = workbook.CreateCellStyle();//创建单元格样式  
                style.SetFont(font);//设置单元格样式中的字体样式  
                cellHyperlink.CellStyle = style;//为单元格设置显示样式  
            }

            using (FileStream fs = File.Create(filePath))
            {
                workbook.Write(fs);
                workbook.Close();
                fs.Close();
            }

            MessageBox.Show("转换完成");
        }

        private void SaveFiles(string dirPath)
        {
            DataTable dt = (DataTable)this.dataGridView1.DataSource;
            foreach (DataRow dr in dt.Rows)
            {
                //如果文件名为空则略过
                if (String.IsNullOrEmpty(dr[0].ToString().Trim()))
                    continue;

                string filepath = dirPath + '/' + dr[0] + ".txt";
                if (File.Exists(filepath))
                {
                    MessageBox.Show("文件已存在:" + filepath);
                    return;
                }
                FileStream fs = File.Create(filepath);

                byte[] content = System.Text.Encoding.Default.GetBytes(dr[1].ToString());
                fs.Write(content, 0, content.Length);

                fs.Close();
            }

            MessageBox.Show("转换完成");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = true;
            if (open.ShowDialog() == DialogResult.OK)
            {
                string[] files = open.FileNames;

                //读取Text
                LoadTexts(files);
            }

            isFromExcel = false;
            isFromTexts = true;
            isFromAny = false;
        }

        private void LoadTexts(string[] files)
        {

            FileStream fs;
            dt.Clear();
            foreach (string file in files)
            {
                fs = File.OpenRead(file);
                DataRow dr = dt.NewRow();
                dr[0] = file.Substring(file.LastIndexOf('\\') + 1, file.LastIndexOf('.') - file.LastIndexOf('\\') - 1);

                byte[] content = new byte[fs.Length];
                fs.Read(content, 0, (int)fs.Length);
                dr[1] = Encoding.Default.GetString(content);

                dt.Rows.Add(dr);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = true;
            if (open.ShowDialog() == DialogResult.OK)
            {
                string[] files = open.FileNames;

                //读取文件
                LoadAnyFiles(files);
            }

            isFromExcel = false;
            isFromTexts = false;
            isFromAny = true;
        }

        private void LoadAnyFiles(string[] files)
        {
            FileStream fs;
            dt.Clear();
            foreach (string file in files)
            {
                fs = File.OpenRead(file);
                DataRow dr = dt.NewRow();
                dr[0] = file.Substring(file.LastIndexOf('\\') + 1);

                dt.Rows.Add(dr);
            }
        }
    }
}
