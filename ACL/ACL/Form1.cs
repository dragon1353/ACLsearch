namespace ACL
{
    using System.Diagnostics;
    using Excel = Microsoft.Office.Interop.Excel;
    using PowerShell = System.Management.Automation.PowerShell;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public bool itembox = true;//初始化寫入combobox清單
        private int item = -1;//判斷伺服器是否有改

        private void button2_Click(object sender, EventArgs e)
        {
            Close();//關閉程式
        }
        private void button3_Click(object sender, EventArgs e)
        {
            int x = 3;//起始欄位資料從3開始
            int y = 0;//流水號用
            string? ID, acl;//帶出EXCEL字串資料
            dataGridView1.Rows.Clear();//初始化資料表
            string pathdata = @"C:\ACL\data.csv";//給路徑
            //確認是否有data資料存在
            if (System.IO.File.Exists(pathdata))
            {
                //讀取EXCEL
                Excel.Application app = new();
                Excel.Workbook wb = app.Workbooks.Open(pathdata);
                //取得工作表
                try
                {
                    Cursor.Current = Cursors.WaitCursor;//滑鼠loading
                    bool IDcheck, ACLcheck;//資料比對是否有涵蓋
                    //初始化欄位數值
                    Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                    Excel.Range arangID = wst.get_Range("D" + x);
                    Excel.Range arang = wst.get_Range("A" + x);
                    //取得欄位數值
                    for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                    {
                        ID = arangID.Text;
                        acl = arang.Text;
#pragma warning disable CS8604 // 可能有 Null 參考引數。
                        IDcheck = ID.Contains(textBox2.Text, StringComparison.OrdinalIgnoreCase);
                        ACLcheck = acl.Contains(comboBox1.Text, StringComparison.OrdinalIgnoreCase);
#pragma warning restore CS8604 // 可能有 Null 參考引數。
                        //如果使用者欄位有值
                        if (IDcheck && textBox2.Text != null && comboBox1.Text == null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //如果資料夾欄位有值
                        else if (ACLcheck && textBox2.Text == null && comboBox1.Text != null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //兩個欄位都有值
                        else if (ACLcheck && IDcheck && textBox2.Text != null && comboBox1.Text != null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //留白
                        else if (textBox1.Text == null && textBox2.Text == null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        else
                        {
                        }
                        //寫入所有權限
                        if (itembox)
                        {
                            //第一筆先無條件加入
                            if (comboBox1.Items.Count == 0)
                            {
                                comboBox1.Items.Add(acl);
                            }
                            //第二筆之後開始比對
                            else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                comboBox1.Items.Add(acl);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                itembox = false; //欄位寫入完畢後不再做動
                //關閉EXCEL後台執行緒 殺掉空白名稱EXCEL執行緒 
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    if (p.MainWindowTitle == "")
                        p.Kill();
                }
                MessageBox.Show("已完成查詢");
            }
            else
            {
                MessageBox.Show("無檔案可查詢");
            }
            Cursor.Current = Cursors.Default;//滑鼠恢復
        }
        //依據選擇伺服器來判斷要執行哪個ps1的function
        private static void Getfile(string path)
        {
            try
            {
                File.GetAttributes(path);
                string cmd = Path.Combine(Directory.GetCurrentDirectory(), path);
                Cursor.Current = Cursors.WaitCursor;//滑鼠loading
                var process = new Process();
                process.StartInfo.FileName = @"Powershell.exe";
                process.StartInfo.Arguments = cmd;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                process.WaitForExit(10000);
                process.Kill();

            }
            catch (Exception)
            {
                MessageBox.Show("powershell可能執行失敗或路徑異常");
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            int x = 3;//起始欄位資料從3開始
            int y = 0;//流水號用
            string? ID, acl;//帶出EXCEL字串資料
            dataGridView1.Rows.Clear();//初始化資料表
            string pathdata = @"C:\ACL\data.csv";//給路徑
            //如果切換伺服器的話就把清單清掉
            if (item != comboBox2.SelectedIndex && comboBox2.SelectedIndex != 0)
            {
                item = comboBox2.SelectedIndex;
                itembox = true;
                comboBox1.Items.Clear();
                comboBox1.Text = "";

            }
            //生成資料表data
            try
            {
                string path = "";
                switch (comboBox2.Text)
                {
                    case "192.168.1.10":
                        path = @"C:\ACL\NETHS1-G.ps1";
                        Getfile(path);
                        break;
                    case "192.168.1.110":
                        path = @"C:\ACL\NETHS1-F.ps1";
                        Getfile(path);
                        break;
                    default:
                        MessageBox.Show("請選擇伺服器");
                        break;
                }
                //先判斷有沒有生成檔案
            }
            catch (Exception)
            {
                MessageBox.Show("選擇伺服器時造成意外中斷");
            }
            if (System.IO.File.Exists(pathdata) && (comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 2))
            {
                //讀取EXCEL
                Excel.Application app = new();
                Excel.Workbook wb = app.Workbooks.Open(pathdata);
                //取得工作表
                try
                {
                    bool IDcheck, ACLcheck;//資料比對是否有涵蓋
                    //初始化欄位數值
                    Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                    Excel.Range arangID = wst.get_Range("D" + x);
                    Excel.Range arang = wst.get_Range("A" + x);
                    //取得欄位數值
                    for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                    {
                        ID = arangID.Text;
                        acl = arang.Text;
#pragma warning disable CS8604 // 可能有 Null 參考引數。
                        IDcheck = ID.Contains(textBox2.Text);
                        ACLcheck = acl.Contains(comboBox1.Text);
#pragma warning restore CS8604 // 可能有 Null 參考引數。
                        //寫入資料到資料表
                        dataGridView1.Rows.Add(y, ID, acl);
                        y++;
                        //寫入所有權限
                        if (itembox)
                        {
                            //第一筆先無條件加入
                            if (comboBox1.Items.Count == 0)
                            {
                                comboBox1.Items.Add(acl);
                            }
                            //第二筆之後開始比對
                            else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                comboBox1.Items.Add(acl);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                itembox = false;//欄位寫入完畢後不再做動
                //關閉EXCEL後台執行緒 殺掉空白名稱EXCEL執行緒 
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    if (p.MainWindowTitle == "")
                        p.Kill();
                }
                MessageBox.Show("已完成資料索引");
                Cursor.Current = Cursors.Default;//滑鼠恢復
            }
            else
            {
                MessageBox.Show("無檔案可查詢");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //初始化資料表欄位
            dataGridView1.ColumnCount = 3;
            this.dataGridView1.Columns[0].HeaderCell.Value = "流水號";
            this.dataGridView1.Columns[1].HeaderCell.Value = "使用者工號";
            this.dataGridView1.Columns[2].HeaderCell.Value = "資料夾權限路徑";
            this.dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridView1.Columns[2].Width = dataGridView1.Width - dataGridView1.Columns[0].Width - dataGridView1.Columns[1].Width;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\ACL\return.csv";//給路徑
            //新增EXCEL,有舊資料就強制覆蓋
            Excel.Application app = new();
            Excel._Workbook wb;
            wb = app.Workbooks.Add();
            Excel._Worksheet ws = wb.Sheets[1];
            try
            {
                ws.Cells[1, 1] = "流水號";
                ws.Cells[1, 2] = "使用者工號";
                ws.Cells[1, 3] = "資料夾權限路徑";
                //輸出清單資料到EXCEL
                for (int countA = 0; countA < dataGridView1.ColumnCount; countA++)
                {
                    for (int countB = 0; countB < dataGridView1.RowCount; countB++)
                    {
                        ws.Cells[countB + 2, countA + 1] = dataGridView1[countA, countB].Value;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            wb.SaveCopyAs(path);//存檔
            MessageBox.Show("已完成資料匯出");
            //關閉EXCEL後台執行緒 殺掉空白名稱EXCEL執行緒 
            Process[] procs = Process.GetProcessesByName("EXCEL");
            foreach (Process p in procs)
            {
                if (p.MainWindowTitle == "")
                    p.Kill();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //條件輸入後可直接ENTER查詢
            if (e.KeyCode == Keys.Enter)
            {
                button3.Focus();
                button3_Click(sender, e);
                textBox2.Focus();
            }

        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //條件輸入後可直接ENTER查詢
            if (e.KeyCode == Keys.Enter)
            {
                button3.Focus();
                button3_Click(sender, e);
                comboBox1.Focus();
            }
        }

        private void dataGridView1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Columns[2].Width = dataGridView1.Width - dataGridView1.Columns[0].Width - dataGridView1.Columns[1].Width;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int x = 3;//起始欄位資料從3開始
            int y = 0;//流水號用
            string? ID, acl;//帶出EXCEL字串資料
            string pathdata = @"C:\ACL\data.csv";//給路徑
                                                 //如果切換伺服器的話就把清單清掉
            if (item != comboBox2.SelectedIndex && comboBox2.SelectedIndex != 0)
            {
                dataGridView1.Rows.Clear();//初始化資料表
                item = comboBox2.SelectedIndex;
                itembox = true;
                comboBox1.Items.Clear();
                comboBox1.Text = "";
                //生成資料表data
                try
                {
                    string path = "";
                    switch (comboBox2.Text)
                    {
                        case "192.168.1.10":
                            path = @"C:\ACL\NETHS1-G.ps1";
                            Getfile(path);
                            break;
                        case "192.168.1.110":
                            path = @"C:\ACL\NETHS1-F.ps1";
                            Getfile(path);
                            break;
                        default:
                            MessageBox.Show("請選擇伺服器");
                            break;
                    }
                    //先判斷有沒有生成檔案
                }
                catch (Exception)
                {
                    MessageBox.Show("選擇伺服器時造成意外中斷");
                }
                if (System.IO.File.Exists(pathdata) && (comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 2))
                {
                    //讀取EXCEL
                    Excel.Application app = new();
                    Excel.Workbook wb = app.Workbooks.Open(pathdata);
                    //取得工作表
                    try
                    {
                        bool IDcheck, ACLcheck;//資料比對是否有涵蓋
                                               //初始化欄位數值
                        Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                        Excel.Range arangID = wst.get_Range("D" + x);
                        Excel.Range arang = wst.get_Range("A" + x);
                        //取得欄位數值
                        for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                        {
                            ID = arangID.Text;
                            acl = arang.Text;
#pragma warning disable CS8604 // 可能有 Null 參考引數。
                            IDcheck = ID.Contains(textBox2.Text);
                            ACLcheck = acl.Contains(comboBox1.Text);
#pragma warning restore CS8604 // 可能有 Null 參考引數。
                            //寫入資料到資料表
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                            //寫入所有權限
                            if (itembox)
                            {
                                //第一筆先無條件加入
                                if (comboBox1.Items.Count == 0)
                                {
                                    comboBox1.Items.Add(acl);
                                }
                                //第二筆之後開始比對
                                else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                    comboBox1.Items.Add(acl);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    itembox = false;//欄位寫入完畢後不再做動
                                    //關閉EXCEL後台執行緒 殺掉空白名稱EXCEL執行緒 
                    Process[] procs = Process.GetProcessesByName("EXCEL");
                    foreach (Process p in procs)
                    {
                        if (p.MainWindowTitle == "")
                            p.Kill();
                    }
                    MessageBox.Show("已完成資料索引");
                }
                else
                {
                    MessageBox.Show("無檔案可查詢");
                }
            }
            else
            {
                if (comboBox2.Text == "請選擇伺服器")
                {
                    MessageBox.Show("請選擇伺服器");
                }
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                try
                {
                    string url = "";
                    dataGridView1.Focus();
#pragma warning disable CS8600 // 正在將 Null 常值或可能的 Null 值轉換為不可為 Null 的型別。
                    url = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
#pragma warning restore CS8600 // 正在將 Null 常值或可能的 Null 值轉換為不可為 Null 的型別。
                    Process.Start("Explorer.exe", $"/e, {url}");
                }
                catch (NullReferenceException)
                {

                    MessageBox.Show("無資料夾或資料");
                }
            }
            else if (e.ColumnIndex == 1)
            {
                try
                {
                    string? keyword;
                    dataGridView1.Focus();
                    keyword = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
#pragma warning disable CS8602 // 可能 null 參考的取值 (dereference)。
                    keyword = keyword.Replace(@"DOMAIN\", "");
#pragma warning restore CS8602 // 可能 null 參考的取值 (dereference)。
                    Form2 form2 = new(keyword);
                    form2.ShowDialog(this);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("無群組或資料");
                }
            }
        }
    }
}