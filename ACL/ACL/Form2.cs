using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace ACL
{
    public partial class Form2 : Form
    {
        public string? keyword;
        public Form2(string keyword)
        {
            InitializeComponent();
            this.keyword = keyword;
            //初始化資料表欄位
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].HeaderCell.Value = "流水號";
            dataGridView1.Columns[1].HeaderCell.Value = "使用者群組";
            dataGridView1.Columns[2].HeaderCell.Value = "使用者帳號";
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ADsearch();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Close();
        }
        //執行寫入EXCEL資料到清單
        public void ADsearch()
        {
            string sub;//寫入整行工號資料比對
            string pathdata = @"C:\ACL\test.csv", pathwrite = @"C:\ACL\test.bat";//路徑
            int number = 1, cou = 9, substr = 0;//流水號、欄位9開始讀取資料、工號字串數
            //先生成dos指令bat檔案
            FileStream fs = File.Create(pathwrite);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(@"net group """ + keyword + @""" /domain > " + pathdata);
            sw.Flush();
            sw.Close();
            if (File.Exists(pathwrite))
            {
                //再使用dos指令將檔案輸出成EXCEL
                Process pathread = new Process();
                pathread.StartInfo.FileName = pathwrite;
                pathread.StartInfo.UseShellExecute = false;
                pathread.StartInfo.CreateNoWindow = true;
                pathread.Start();
                pathread.WaitForExit(3000);
                pathread.CloseMainWindow();
                //讀取EXCEL
                Excel.Application app = new();
                Excel.Workbook wb = app.Workbooks.Open(pathdata);
                //取得工作表
                try
                {
                    Cursor.Current = Cursors.WaitCursor;//滑鼠loading
                    string str = "";//記錄工號
                    Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["test"];
                    Excel.Range ID = wst.get_Range("A" + cou);//工號
                    for (cou = 9; ID.Value2 != null; cou++, ID = wst.get_Range("A" + cou))
                    {
                        sub = ID.Value2;
                        //整行讀入如果非空白及非null就一個字一個字載入
                        for (substr = 0; substr < sub.Length; substr++)
                        {
                            if (sub.Substring(substr, 1) is not " ")
                            {
                                str += sub.Substring(substr, 1);
                            }
                            else
                            {
                                //如果字串讀完有資料就寫入資料表
                                if (str != "")
                                {
                                    dataGridView1.Rows.Add(number, keyword, str);
                                    number++;//流水號增加
                                    str = "";//重置寫入字串
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("無群組或人員資料");
                }
                Cursor.Current = Cursors.Default;//滑鼠恢復
                //關閉EXCEL後台執行緒 殺掉空白名稱EXCEL執行緒 
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    if (p.MainWindowTitle == "")
                        p.Kill();
                }
            }
            else 
            {
                MessageBox.Show("無生成檔案");
            }
        }
    }
}
