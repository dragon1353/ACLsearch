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
        public bool itembox = true;//��l�Ƽg�Jcombobox�M��
        private int item = -1;//�P�_���A���O�_����

        private void button2_Click(object sender, EventArgs e)
        {
            Close();//�����{��
        }
        private void button3_Click(object sender, EventArgs e)
        {
            int x = 3;//�_�l����Ʊq3�}�l
            int y = 0;//�y������
            string? ID, acl;//�a�XEXCEL�r����
            dataGridView1.Rows.Clear();//��l�Ƹ�ƪ�
            string pathdata = @"C:\ACL\data.csv";//�����|
            //�T�{�O�_��data��Ʀs�b
            if (System.IO.File.Exists(pathdata))
            {
                //Ū��EXCEL
                Excel.Application app = new();
                Excel.Workbook wb = app.Workbooks.Open(pathdata);
                //���o�u�@��
                try
                {
                    Cursor.Current = Cursors.WaitCursor;//�ƹ�loading
                    bool IDcheck, ACLcheck;//��Ƥ��O�_���[�\
                    //��l�����ƭ�
                    Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                    Excel.Range arangID = wst.get_Range("D" + x);
                    Excel.Range arang = wst.get_Range("A" + x);
                    //���o���ƭ�
                    for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                    {
                        ID = arangID.Text;
                        acl = arang.Text;
#pragma warning disable CS8604 // �i�঳ Null �ѦҤ޼ơC
                        IDcheck = ID.Contains(textBox2.Text, StringComparison.OrdinalIgnoreCase);
                        ACLcheck = acl.Contains(comboBox1.Text, StringComparison.OrdinalIgnoreCase);
#pragma warning restore CS8604 // �i�঳ Null �ѦҤ޼ơC
                        //�p�G�ϥΪ���즳��
                        if (IDcheck && textBox2.Text != null && comboBox1.Text == null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //�p�G��Ƨ���즳��
                        else if (ACLcheck && textBox2.Text == null && comboBox1.Text != null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //�����쳣����
                        else if (ACLcheck && IDcheck && textBox2.Text != null && comboBox1.Text != null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        //�d��
                        else if (textBox1.Text == null && textBox2.Text == null)
                        {
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                        }
                        else
                        {
                        }
                        //�g�J�Ҧ��v��
                        if (itembox)
                        {
                            //�Ĥ@�����L����[�J
                            if (comboBox1.Items.Count == 0)
                            {
                                comboBox1.Items.Add(acl);
                            }
                            //�ĤG������}�l���
                            else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                comboBox1.Items.Add(acl);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                itembox = false; //���g�J�����ᤣ�A����
                //����EXCEL��x����� �����ťզW��EXCEL����� 
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    if (p.MainWindowTitle == "")
                        p.Kill();
                }
                MessageBox.Show("�w�����d��");
            }
            else
            {
                MessageBox.Show("�L�ɮץi�d��");
            }
            Cursor.Current = Cursors.Default;//�ƹ���_
        }
        //�̾ڿ�ܦ��A���ӧP�_�n�������ps1��function
        private static void Getfile(string path)
        {
            try
            {
                File.GetAttributes(path);
                string cmd = Path.Combine(Directory.GetCurrentDirectory(), path);
                Cursor.Current = Cursors.WaitCursor;//�ƹ�loading
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
                MessageBox.Show("powershell�i����楢�ѩθ��|���`");
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            int x = 3;//�_�l����Ʊq3�}�l
            int y = 0;//�y������
            string? ID, acl;//�a�XEXCEL�r����
            dataGridView1.Rows.Clear();//��l�Ƹ�ƪ�
            string pathdata = @"C:\ACL\data.csv";//�����|
            //�p�G�������A�����ܴN��M��M��
            if (item != comboBox2.SelectedIndex && comboBox2.SelectedIndex != 0)
            {
                item = comboBox2.SelectedIndex;
                itembox = true;
                comboBox1.Items.Clear();
                comboBox1.Text = "";

            }
            //�ͦ���ƪ�data
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
                        MessageBox.Show("�п�ܦ��A��");
                        break;
                }
                //���P�_���S���ͦ��ɮ�
            }
            catch (Exception)
            {
                MessageBox.Show("��ܦ��A���ɳy���N�~���_");
            }
            if (System.IO.File.Exists(pathdata) && (comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 2))
            {
                //Ū��EXCEL
                Excel.Application app = new();
                Excel.Workbook wb = app.Workbooks.Open(pathdata);
                //���o�u�@��
                try
                {
                    bool IDcheck, ACLcheck;//��Ƥ��O�_���[�\
                    //��l�����ƭ�
                    Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                    Excel.Range arangID = wst.get_Range("D" + x);
                    Excel.Range arang = wst.get_Range("A" + x);
                    //���o���ƭ�
                    for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                    {
                        ID = arangID.Text;
                        acl = arang.Text;
#pragma warning disable CS8604 // �i�঳ Null �ѦҤ޼ơC
                        IDcheck = ID.Contains(textBox2.Text);
                        ACLcheck = acl.Contains(comboBox1.Text);
#pragma warning restore CS8604 // �i�঳ Null �ѦҤ޼ơC
                        //�g�J��ƨ��ƪ�
                        dataGridView1.Rows.Add(y, ID, acl);
                        y++;
                        //�g�J�Ҧ��v��
                        if (itembox)
                        {
                            //�Ĥ@�����L����[�J
                            if (comboBox1.Items.Count == 0)
                            {
                                comboBox1.Items.Add(acl);
                            }
                            //�ĤG������}�l���
                            else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                comboBox1.Items.Add(acl);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                itembox = false;//���g�J�����ᤣ�A����
                //����EXCEL��x����� �����ťզW��EXCEL����� 
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    if (p.MainWindowTitle == "")
                        p.Kill();
                }
                MessageBox.Show("�w������Ư���");
                Cursor.Current = Cursors.Default;//�ƹ���_
            }
            else
            {
                MessageBox.Show("�L�ɮץi�d��");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //��l�Ƹ�ƪ����
            dataGridView1.ColumnCount = 3;
            this.dataGridView1.Columns[0].HeaderCell.Value = "�y����";
            this.dataGridView1.Columns[1].HeaderCell.Value = "�ϥΪ̤u��";
            this.dataGridView1.Columns[2].HeaderCell.Value = "��Ƨ��v�����|";
            this.dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridView1.Columns[2].Width = dataGridView1.Width - dataGridView1.Columns[0].Width - dataGridView1.Columns[1].Width;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\ACL\return.csv";//�����|
            //�s�WEXCEL,���¸�ƴN�j���л\
            Excel.Application app = new();
            Excel._Workbook wb;
            wb = app.Workbooks.Add();
            Excel._Worksheet ws = wb.Sheets[1];
            try
            {
                ws.Cells[1, 1] = "�y����";
                ws.Cells[1, 2] = "�ϥΪ̤u��";
                ws.Cells[1, 3] = "��Ƨ��v�����|";
                //��X�M���ƨ�EXCEL
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
            wb.SaveCopyAs(path);//�s��
            MessageBox.Show("�w������ƶץX");
            //����EXCEL��x����� �����ťզW��EXCEL����� 
            Process[] procs = Process.GetProcessesByName("EXCEL");
            foreach (Process p in procs)
            {
                if (p.MainWindowTitle == "")
                    p.Kill();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //�����J��i����ENTER�d��
            if (e.KeyCode == Keys.Enter)
            {
                button3.Focus();
                button3_Click(sender, e);
                textBox2.Focus();
            }

        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //�����J��i����ENTER�d��
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
            int x = 3;//�_�l����Ʊq3�}�l
            int y = 0;//�y������
            string? ID, acl;//�a�XEXCEL�r����
            string pathdata = @"C:\ACL\data.csv";//�����|
                                                 //�p�G�������A�����ܴN��M��M��
            if (item != comboBox2.SelectedIndex && comboBox2.SelectedIndex != 0)
            {
                dataGridView1.Rows.Clear();//��l�Ƹ�ƪ�
                item = comboBox2.SelectedIndex;
                itembox = true;
                comboBox1.Items.Clear();
                comboBox1.Text = "";
                //�ͦ���ƪ�data
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
                            MessageBox.Show("�п�ܦ��A��");
                            break;
                    }
                    //���P�_���S���ͦ��ɮ�
                }
                catch (Exception)
                {
                    MessageBox.Show("��ܦ��A���ɳy���N�~���_");
                }
                if (System.IO.File.Exists(pathdata) && (comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 2))
                {
                    //Ū��EXCEL
                    Excel.Application app = new();
                    Excel.Workbook wb = app.Workbooks.Open(pathdata);
                    //���o�u�@��
                    try
                    {
                        bool IDcheck, ACLcheck;//��Ƥ��O�_���[�\
                                               //��l�����ƭ�
                        Excel._Worksheet wst = (Excel._Worksheet)wb.Worksheets["data"];
                        Excel.Range arangID = wst.get_Range("D" + x);
                        Excel.Range arang = wst.get_Range("A" + x);
                        //���o���ƭ�
                        for (x = 3, y = 1; arangID.Value2 != null; x++, arangID = wst.get_Range("D" + x), arang = wst.get_Range("A" + x))
                        {
                            ID = arangID.Text;
                            acl = arang.Text;
#pragma warning disable CS8604 // �i�঳ Null �ѦҤ޼ơC
                            IDcheck = ID.Contains(textBox2.Text);
                            ACLcheck = acl.Contains(comboBox1.Text);
#pragma warning restore CS8604 // �i�঳ Null �ѦҤ޼ơC
                            //�g�J��ƨ��ƪ�
                            dataGridView1.Rows.Add(y, ID, acl);
                            y++;
                            //�g�J�Ҧ��v��
                            if (itembox)
                            {
                                //�Ĥ@�����L����[�J
                                if (comboBox1.Items.Count == 0)
                                {
                                    comboBox1.Items.Add(acl);
                                }
                                //�ĤG������}�l���
                                else if (acl != comboBox1.Items[comboBox1.Items.Count - 1].ToString())
                                    comboBox1.Items.Add(acl);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    itembox = false;//���g�J�����ᤣ�A����
                                    //����EXCEL��x����� �����ťզW��EXCEL����� 
                    Process[] procs = Process.GetProcessesByName("EXCEL");
                    foreach (Process p in procs)
                    {
                        if (p.MainWindowTitle == "")
                            p.Kill();
                    }
                    MessageBox.Show("�w������Ư���");
                }
                else
                {
                    MessageBox.Show("�L�ɮץi�d��");
                }
            }
            else
            {
                if (comboBox2.Text == "�п�ܦ��A��")
                {
                    MessageBox.Show("�п�ܦ��A��");
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
#pragma warning disable CS8600 // ���b�N Null �`�ȩΥi�઺ Null ���ഫ�����i�� Null �����O�C
                    url = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
#pragma warning restore CS8600 // ���b�N Null �`�ȩΥi�઺ Null ���ഫ�����i�� Null �����O�C
                    Process.Start("Explorer.exe", $"/e, {url}");
                }
                catch (NullReferenceException)
                {

                    MessageBox.Show("�L��Ƨ��θ��");
                }
            }
            else if (e.ColumnIndex == 1)
            {
                try
                {
                    string? keyword;
                    dataGridView1.Focus();
                    keyword = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
#pragma warning disable CS8602 // �i�� null �ѦҪ����� (dereference)�C
                    keyword = keyword.Replace(@"DOMAIN\", "");
#pragma warning restore CS8602 // �i�� null �ѦҪ����� (dereference)�C
                    Form2 form2 = new(keyword);
                    form2.ShowDialog(this);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("�L�s�թθ��");
                }
            }
        }
    }
}