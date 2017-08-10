using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace EAT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
          
            open.Filter= "Excel 工作簿|*.xlsx|Excel 97-2003 工作簿|*.xls";

            if (open.ShowDialog() == DialogResult.OK)
            {
                string filePath = open.FileName;

                //解析excel
                LoadExcel(filePath);
            }
            
        }

        private void LoadExcel(string filePath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Sheets sheets;
            Excel.Workbook workbook = null;
            object oMissiong = Missing.Value;
            DataTable dt = new DataTable();

            workbook = app.Workbooks.Open(filePath, oMissiong, true, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
            //将数据读入到DataTable中——Start    

            sheets = workbook.Worksheets;
            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);//读取第一张表  
            if (worksheet == null)
            {
                MessageBox.Show("Excel表中没有数据");
                return;
            }


            string cellContent;
            string cellName;
            int iRowCount = worksheet.UsedRange.Rows.Count;
            int iColCount = worksheet.UsedRange.Columns.Count;
            Excel.Range range;

       
            this.dataGridView1.DataSource = dt; 
            dt.Columns.Add(new DataColumn("文件名"));
            dt.Columns.Add(new DataColumn("内容"));
            this.dataGridView1.Columns[1].Width = 300;


            for (int i = 1; i <= iRowCount; ++i)
            {
                range = (Excel.Range)worksheet.Cells[i, 1];
                cellName = range.Text.ToString();

                range = (Excel.Range)worksheet.Cells[i, 2];
                cellContent = range.Text.ToString();

                DataRow dr = dt.NewRow();
                dr[0] = cellName;
                dr[1] = cellContent;

                dt.Rows.Add(dr);
            }


            //将数据读入到DataTable中——End  
            workbook.Close(false, oMissiong, oMissiong);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            app.Workbooks.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("版权说明: \n\n本软件由中山大学肖鹏博士团队开发\n设计策划：肖鹏，联系方式：cifangyue@163.com\n技术负责：王海明 肖鹏");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                string dirPath = folder.SelectedPath;

                SaveFiles(dirPath);
            }
        }

        private void SaveFiles(string dirPath)
        {
            DataTable dt = (DataTable)this.dataGridView1.DataSource;
            foreach(DataRow dr in dt.Rows)
            {
                string filepath = dirPath + '/' + dr[0]+".txt";
                if (File.Exists(filepath))
                {
                    MessageBox.Show("文件已存在:" + filepath);
                    return;
                }
                FileStream fs = File.Create(filepath);
                
                byte[] content= System.Text.Encoding.Default.GetBytes(dr[1].ToString());
                fs.Write(content, 0, content.Length);

                fs.Close();
            }

            MessageBox.Show("转换完成");
        }
    }
}
