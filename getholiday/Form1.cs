using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace getholiday
{
    public partial class Form1 : Form
    {
        BackgroundWorker worker = new BackgroundWorker();
        BackgroundWorker worker2 = new BackgroundWorker();
        Form2 form2 = new Form2();
        Form3 form3 = new Form3();
        Dictionary<string, int> codedic = new Dictionary<string, int>();
        public class restGroup
        {
            public List<string> groups = new List<string>();
        }
        public class driver
        {
            public string companyNumber = "";
            public string name = "";
            public string workGroup = "";
            public string chagoname= "";
            public string tempRestGroup="";
        }
        public List<driver> drivers;
        public List<restGroup> restGroups;
        public Form1()
        {
            InitializeComponent();
        }
        private void ReleaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        void getworkerlist()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                drivers = null;
                drivers = new List<driver>();


                textBox1.Text = ofd.FileName;

                Excel.Application excelApp = null;
                Excel.Workbook wb = null;
                Excel.Worksheet ws = null;

                try
                {
                    excelApp = new Excel.Application();
                    wb = excelApp.Workbooks.Open(textBox1.Text);
                    ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                    Excel.Range rng = ws.UsedRange;

                    object[,] data = rng.Value;
                    for (int r = 2; r <= data.GetLength(0); r++)
                    {
                        driver ndriver = new driver();
                        ndriver.companyNumber = data[r, 2].ToString();
                        ndriver.name = data[r, 3].ToString();
                        ndriver.workGroup = data[r, 6].ToString();
                        ndriver.chagoname = data[r, 4].ToString();
                        ndriver.tempRestGroup = data[r, 5].ToString();
                        drivers.Add(ndriver);
                    }

                    MessageBox.Show("저장완료");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    ReleaseExcelObject(ws);
                    ReleaseExcelObject(wb);
                    excelApp.Quit();
                    ReleaseExcelObject(excelApp);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                drivers = null;
                drivers = new List<driver>();


                textBox1.Text = ofd.FileName;

                Excel.Application excelApp = null;
                Excel.Workbook wb = null;
                Excel.Worksheet ws = null;

                try
                {
                    excelApp = new Excel.Application();
                    wb = excelApp.Workbooks.Open(textBox1.Text);
                    ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                    Excel.Range rng = ws.UsedRange;

                    object[,] data = rng.Value;
                    for (int r = 2; r <= data.GetLength(0); r++)
                    {
                        driver ndriver = new driver();
                        ndriver.companyNumber = data[r, 2].ToString();
                        ndriver.name = data[r, 3].ToString();
                        ndriver.workGroup = data[r, 6].ToString();
                        ndriver.chagoname = data[r, 4].ToString();
                        ndriver.tempRestGroup = data[r, 5].ToString();
                        drivers.Add(ndriver);
                    }

                    MessageBox.Show("저장완료");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    ReleaseExcelObject(ws);
                    ReleaseExcelObject(wb);
                    excelApp.Quit();
                    ReleaseExcelObject(excelApp);
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if(ofd.ShowDialog() == DialogResult.OK) { 
            textBox2.Text = ofd.FileName;
            restGroups = null;
            restGroups = new List<restGroup>();
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            int days = DateTime.DaysInMonth(int.Parse(dateTimePicker1.Value.ToString("yyyy")), int.Parse(dateTimePicker1.Value.ToString("MM")));
            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(textBox2.Text);
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                Excel.Range rng = ws.UsedRange;
                object[,] data = rng.Value;
                    data.SetValue(1000, 0);
                    data.SetValue(1000, 1);
                    MessageBox.Show(data.GetLength(0).ToString());
                    MessageBox.Show(data.GetLength(1).ToString());
                    for (int r = 5; r <= days+4; r++)
                {
                    restGroup rgroup = new restGroup();
                    char[] sep = { ',' };
                    string Groups = data[r, 11].ToString();
                    string[] temp = Groups.Split(sep);
                    foreach(string s in temp)
                    {
                        rgroup.groups.Add(s);
                    }
                    restGroups.Add(rgroup);
                }
                MessageBox.Show("저장완료");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                excelApp.Quit();
                ReleaseExcelObject(excelApp);
            }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string name = textBox3.Text;
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Server=150.50.78.152;database=tempdb;uid=sa;pwd=2@@51111";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            conn.Open();
            cmd.CommandText = "select id from [handicall].[dbo].[AgentInfo] WHERE name = '"+name+"'";
            int id = Convert.ToInt32(cmd.ExecuteScalar());

            if (textBox4.Text != "" && textBox1.Text !=""&&textBox2.Text!="")
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
                restGroup rgroup = new restGroup();
                char[] sep = { ',' };
                string Groups = textBox4.Text.ToString();
                string[] temp = Groups.Split(sep);
                foreach (string s in temp)
                {
                    rgroup.groups.Add(s);
                }



            for (int i = 0; i < drivers.Count; i++)
            {
                int count = 1;
                driver d = drivers[i];
                    bool firstdaycheck = false;
                if (d.tempRestGroup.Substring(0, 2) != "야간")
                {
                        int driverworkgroupcode = 0;
                        int drivertemprestgroupcode = 0;
                        int chagonamecode = 0;

                        cmd.CommandText = "select cbcode from [handicall].[dbo].[CarBase] WHERE cbname = '" +d.chagoname+ "'";
                        chagonamecode = Convert.ToInt32(cmd.ExecuteScalar());

                        cmd.CommandText = "select code from [handicall].[dbo].[WorkClassName] WHERE name = '" + d.workGroup + "'";
                        driverworkgroupcode = Convert.ToInt32(cmd.ExecuteScalar());

                        cmd.CommandText = "select code from [handicall].[dbo].[WorkClassName] WHERE name = '" + d.tempRestGroup + "'";
                        drivertemprestgroupcode = Convert.ToInt32(cmd.ExecuteScalar());
                        for (int j = 0; j < restGroups.Count; j++)
                    {
                            if (j == 0 && firstdaycheck == false)
                            {
                                for (int t = 0; t < rgroup.groups.Count; t++)
                                {
                                    if (havingalpha(rgroup.groups[t]))
                                    {
                                        if (except_Hyphen_Zero(d.tempRestGroup) == rgroup.groups[t])
                                        {
                                            int temp_returnday = returnday(d.tempRestGroup, -1);
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                              dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode, dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id , "근무조정");
                                            
                                            j = temp_returnday;
                                            firstdaycheck = true;
                                            if (temp_returnday >= restGroups.Count) break;

                                        }


                                    }
                                    else
                                    {
                                        if (onlyNumber(d.tempRestGroup) == rgroup.groups[t])
                                        {
                                            int temp_returnday = returnday(d.tempRestGroup, -1);
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                            j = temp_returnday;
                                            firstdaycheck = true;
                                            if (temp_returnday >= restGroups.Count) break;

                                        }

                                    }


                                    
                                }
                                if(firstdaycheck == false)
                                {
                                    for (int k = 0; k < restGroups[j].groups.Count; k++)
                                    {
                                        if (havingalpha(restGroups[j].groups[k]))
                                        {
                                            if (except_Hyphen_Zero(d.tempRestGroup) == restGroups[j].groups[k])
                                            {
                                                dataGridView1.Rows.Add(d.companyNumber, d.name, d.tempRestGroup, drivertemprestgroupcode, d.chagoname, chagonamecode,
                                                    dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                                dataGridView2.Rows.Add(d.companyNumber, d.name, drivertemprestgroupcode, chagonamecode,
                                                                                        dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                                count++;
                                                int temp_returnday = returnday(d.tempRestGroup, j);
                                                if (temp_returnday+1 <= DateTime.DaysInMonth(int.Parse(dateTimePicker1.Value.ToString("yyyy")), int.Parse(dateTimePicker1.Value.ToString("MM"))))
                                                {
                                                    dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                                       dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                                    dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode, dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"), id,"근무조정");

                                                }
                                                count++;
                                                if (temp_returnday >= restGroups.Count) break;
                                                j = temp_returnday;
                                            }
                                        }
                                        else
                                        {
                                            if (onlyNumber(d.tempRestGroup) == restGroups[j].groups[k])
                                            {
                                                dataGridView1.Rows.Add(d.companyNumber, d.name, d.tempRestGroup, drivertemprestgroupcode, d.chagoname, chagonamecode,
                                                   dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                                dataGridView2.Rows.Add(d.companyNumber, d.name, drivertemprestgroupcode, chagonamecode,
                                                                                  dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                                count++;
                                                int temp_returnday = returnday(d.tempRestGroup, j);
                                                if (temp_returnday+1 <= DateTime.DaysInMonth(int.Parse(dateTimePicker1.Value.ToString("yyyy")), int.Parse(dateTimePicker1.Value.ToString("MM"))))
                                                {
                                                    dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                                       dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                                    dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode,
                                                       dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                                }
                                                count++;
                                                if (temp_returnday >= restGroups.Count) break;
                                                j = temp_returnday;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int k = 0; k < restGroups[j].groups.Count; k++)
                                {
                                    if (havingalpha(restGroups[j].groups[k]))
                                    {
                                        if (except_Hyphen_Zero(d.tempRestGroup) == restGroups[j].groups[k])
                                        {
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.tempRestGroup, drivertemprestgroupcode, d.chagoname, chagonamecode,
                                                dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, drivertemprestgroupcode, chagonamecode,
                                                                                    dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                            count++;
                                            int temp_returnday = returnday(d.tempRestGroup, j);
                                            if(temp_returnday+1 <= DateTime.DaysInMonth(int.Parse(dateTimePicker1.Value.ToString("yyyy")), int.Parse(dateTimePicker1.Value.ToString("MM")))) { 
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode,
                                                                                   dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");

                                            }
                                            count++;
                                            j = temp_returnday;
                                            if (temp_returnday >= restGroups.Count) break;

                                        }
                                    }
                                    else
                                    {
                                        if (onlyNumber(d.tempRestGroup) == restGroups[j].groups[k])
                                        {
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.tempRestGroup, drivertemprestgroupcode, d.chagoname, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, drivertemprestgroupcode, chagonamecode,
                                                                              dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (j + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                            count++;
                                            int temp_returnday = returnday(d.tempRestGroup, j);
                                            if(temp_returnday+1<= DateTime.DaysInMonth(int.Parse(dateTimePicker1.Value.ToString("yyyy")), int.Parse(dateTimePicker1.Value.ToString("MM")))) { 
                                            dataGridView1.Rows.Add(d.companyNumber, d.name, d.workGroup, driverworkgroupcode, d.chagoname, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), textBox3.Text, DateTime.Now.ToString("yyyy-MM-dd"), "근무조정");
                                            dataGridView2.Rows.Add(d.companyNumber, d.name, driverworkgroupcode, chagonamecode,
                                               dateTimePicker1.Value.ToString("yyyy-MM") + "-" + (temp_returnday + 1).ToString(), 1, DateTime.Now.ToString("yyyy-MM-dd"),id, "근무조정");
                                            }
                                            count++;
                                            j = temp_returnday;
                                            if (temp_returnday >= restGroups.Count) break;

                                        }
                                    }
                                }
                            }
                    }
                }
            }
            button4.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
        }
            else
            {
                MessageBox.Show("빈칸을 모두 채워주세요!");
            }
        }
        private int returnday(string raw_workday, int j)
        {
            int num = j;
            if (num + 1 >= restGroups.Count) return (num + 1);
            while(num+1< restGroups.Count)
            {
                bool check = false;
                num++;
                for (int i = 0; i < restGroups[num].groups.Count; i++)
                {
                    if (havingalpha(restGroups[num].groups[i]))
                    {
                        if(except_Hyphen_Zero(raw_workday) == restGroups[num].groups[i])
                        {
                            check = true;
                            break;
                        }
                    }
                    else
                    {
                        if(onlyNumber(raw_workday) == restGroups[num].groups[i])
                        {
                            check = true;
                            break;
                        }
                    }
                }
                if (check == false) return num;
                if (num + 1 == restGroups.Count) return num + 1;
            }
            return num;
        }
        private string except_Hyphen_Zero(string s)
        {
            string newString = "";
            char[] token = { '-' };
            string[] temp = s.Split(token);
            newString = int.Parse(temp[0]).ToString() + temp[1];
            return newString;
        }
        private string onlyNumber(string s)
        {
            string newString = "";
            char[] token = { '-' };
            string[] temp = s.Split(token);
            newString = int.Parse(temp[0]).ToString(); 
            return newString;
        }
        private bool havingalpha(string s)
        {
            if (s.Contains('A') || s.Contains('B') || s.Contains('C') || s.Contains('D') || s.Contains('E') || s.Contains('F') || s.Contains('G') || s.Contains('H') || s.Contains('I') || s.Contains('J') || s.Contains('K') || s.Contains('L') || s.Contains('M') || s.Contains('N') || s.Contains('O') || s.Contains('P') || s.Contains('Q') || s.Contains('R') || s.Contains('S') || s.Contains('T') || s.Contains('U') || s.Contains('V') || s.Contains('W') || s.Contains('X') || s.Contains('Y') || s.Contains('Z')) return true;
            else return false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView2.AllowUserToAddRows = false;
            dataGridView1.AllowUserToAddRows = false;
            drivers = new List<driver>();
            restGroups = new List<restGroup>();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM";
            /*codedic.Add("도봉산차고지", 1);
            codedic.Add("수락산주차장", 3);
            codedic.Add("화랑대주차장", 5);
            codedic.Add("면목홈플러스", 6);
            codedic.Add("서울의료원", 7);
            codedic.Add("어린이대공원",8);
            codedic.Add("용답주차장",9);
            codedic.Add("월드컵경기장",12);
            codedic.Add("상암주차장",13);
            codedic.Add("시립은평병원",16);
            codedic.Add("천호주차장",17);
            codedic.Add("신천유수지",18);
            codedic.Add("목동주차장",25);
            codedic.Add("개화산주차장",30);
            codedic.Add("동대문구청",31);
            codedic.Add("수서주차장", 32);
            codedic.Add("창동차고지", 34);
            codedic.Add("으뜸공원차고지", 35);
            codedic.Add("천왕환승", 36);
            codedic.Add("구파발환승", 37);
            codedic.Add("대림운동장", 38);
            codedic.Add("북서울꿈의숲", 40);
            codedic.Add("신길환승", 41);
            codedic.Add("동작주차공원", 43);
            codedic.Add("한우리문화센터", 45);
            codedic.Add("용산차고지", 46);
            codedic.Add("영등포구청", 59);
            codedic.Add("신림공영주차장", 60);
            codedic.Add("청암공영주차장", 61);
            codedic.Add("종묘지하주차장", 62);
            codedic.Add("사당노외주차장", 63);
            codedic.Add("남부여성발전센터", 64);
            codedic.Add("아리랑차고지", 65);
            codedic.Add("마포유수지", 67);
            codedic.Add("서대문구의회", 68);
            codedic.Add("독립공원", 69);
            codedic.Add("훈련원주차장", 70);
            codedic.Add("신방화차고지", 71);
            codedic.Add("가재울차고지", 72);
            codedic.Add("고척차고지", 73);
            codedic.Add("혁신파크차고지", 74);
            codedic.Add("양평교차고지", 75);

            codedic.Add("01-A", 32);
            codedic.Add("01-B", 33);
            codedic.Add("02-A", 34);
            codedic.Add("02-B", 35);
            codedic.Add("03-A", 36);
            codedic.Add("03-B", 37);
            codedic.Add("04-A", 38);
            codedic.Add("04-B", 39);
            codedic.Add("05-A", 40);
            codedic.Add("05-B", 41);
            codedic.Add("06-A", 42);
            codedic.Add("06-B", 43);
            codedic.Add("07-A", 44);
            codedic.Add("07-B", 45);
            codedic.Add("08-A", 46);
            codedic.Add("08-B", 47);
            codedic.Add("09-A", 48);
            codedic.Add("09-B", 49);
            codedic.Add("10-A", 50);
            codedic.Add("10-B", 51);

            codedic.Add("01-C", 136);
            codedic.Add("01-D", 137);
            codedic.Add("01-E", 138);
            codedic.Add("02-C", 139);
            codedic.Add("02-D", 140);
            codedic.Add("02-E", 141);
            codedic.Add("03-C", 142);
            codedic.Add("03-D", 143);
            codedic.Add("03-E", 144);
            codedic.Add("04-C", 145);
            codedic.Add("04-D", 146);
            codedic.Add("04-E", 147);
            codedic.Add("05-C", 148);
            codedic.Add("05-D", 149);
            codedic.Add("05-E", 150);
            codedic.Add("06-C", 151);
            codedic.Add("06-D", 152);
            codedic.Add("06-E", 153);
            codedic.Add("07-C", 154);
            codedic.Add("07-D", 155);
            codedic.Add("07-E", 156);
            codedic.Add("08-C", 157);
            codedic.Add("08-D", 158);
            codedic.Add("08-E", 159);
            codedic.Add("09-C", 160);
            codedic.Add("09-D", 161);
            codedic.Add("09-E", 162);
            codedic.Add("10-C", 163);
            codedic.Add("10-D", 164);
            codedic.Add("10-E", 165);
            codedic.Add("11-A", 166);
            codedic.Add("11-B", 167);
            codedic.Add("11-C", 168);
            codedic.Add("11-D", 169);
            codedic.Add("11-E", 170);
            codedic.Add("12-A", 171);
            codedic.Add("12-B", 172);
            codedic.Add("12-C", 173);
            codedic.Add("12-D", 174);
            codedic.Add("12-E", 175);
            codedic.Add("11-X", 181);
            codedic.Add("13-A", 183);
            codedic.Add("13-B", 184);
            codedic.Add("14-A", 198);
            codedic.Add("14-B", 199);
            codedic.Add("15-A", 201);
            codedic.Add("15-B", 202);
            codedic.Add("16-A", 203);
            codedic.Add("16-B", 204);
            codedic.Add("17-A", 205);
            codedic.Add("17-B", 206);
            codedic.Add("18-A", 207);
            codedic.Add("18-B", 208);
            codedic.Add("19-A", 209);
            codedic.Add("19-B", 215);
            codedic.Add("14-C", 216);
            codedic.Add("14-D", 217);
            codedic.Add("14-E", 218);
            codedic.Add("15-C", 219);
            codedic.Add("15-D", 220);
            codedic.Add("15-E", 221);*/

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1);
        }
        public void ExportToExcel(DataGridView dataGridView)

        {

            Excel.Application excelApplication;

            Excel._Workbook workbook;

            Excel._Worksheet worksheet;



            excelApplication = new Excel.Application();



            excelApplication.Visible = false;



            workbook = excelApplication.Workbooks.Add(Type.Missing);

            worksheet = workbook.ActiveSheet as Excel._Worksheet;





            object[,] headerValueArray = new object[1, dataGridView.ColumnCount];



            for (int i = 0; i < dataGridView.Columns.Count; i++)

            {

                headerValueArray[0, i] = dataGridView.Columns[i].HeaderText;

            }



            Excel.Range startHeaderCell = worksheet.Cells[1, 1];

            Excel.Range endHeaderCell = worksheet.Cells[1, dataGridView.ColumnCount];



            worksheet.get_Range

            (

                startHeaderCell as object,

                endHeaderCell as object

            ).Font.Bold = true;



            worksheet.get_Range

            (

                startHeaderCell as object,

                endHeaderCell as object

            ).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;



            worksheet.get_Range(startHeaderCell as object, endHeaderCell as object).Value2 = headerValueArray;



            object[,] cellValueArray = new object[dataGridView.RowCount, dataGridView.ColumnCount];



            for (int i = 0; i < dataGridView.RowCount; i++)

            {

                for (int j = 0; j < dataGridView.ColumnCount; j++)

                {

                        if (dataGridView.Rows[i].Cells[j].ValueType.Name == "String")

                        {

                            cellValueArray[i, j] = "'" + dataGridView.Rows[i].Cells[j].Value.ToString();

                        }

                        else

                        {

                            cellValueArray[i, j] = dataGridView.Rows[i].Cells[j].Value;

                        }
                    
                }

            }



            Excel.Range startCell = worksheet.Cells[2, 1];

            Excel.Range endCell = worksheet.Cells[dataGridView.RowCount + 1, dataGridView.ColumnCount];



            worksheet.get_Range(startCell as object, endCell as object).Value2 = cellValueArray;



            excelApplication.Visible = true;



            excelApplication.UserControl = true;

        }
        void updateDB()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Server=150.50.78.152;database=tempdb;uid=sa;pwd=2@@51111";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            conn.Open();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Cells[0].Value.ToString() != "")
                {
                    string companynum = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    string name = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    string carclasscode = dataGridView2.Rows[i].Cells[2].Value.ToString();
                    string carbasecode = dataGridView2.Rows[i].Cells[3].Value.ToString();
                    string changeDate = dataGridView2.Rows[i].Cells[4].Value.ToString();
                    string RegDate = dataGridView2.Rows[i].Cells[6].Value.ToString();
                    string id = dataGridView2.Rows[i].Cells[7].Value.ToString();
                    cmd.CommandText = "IF NOT EXISTS(SELECT * FROM [handicall].[dbo].[DriverInfoScheduleChange] WHERE no = " + companynum + " AND name='" + name + "' AND carclasscode= " + carclasscode + " AND changeDate = '" + changeDate + "') BEGIN INSERT INTO [handicall].[dbo].[DriverInfoScheduleChange] values (" + companynum + ",'" + name + "'," + carclasscode + "," + carbasecode + ",'" + changeDate + "',1,'" + RegDate + "'," + id + ",'근무조정') END";
                    cmd.ExecuteNonQuery();
                }
            }
            conn.Close();
        }
        void loadingcomplete()
        {
            form2.Close();
        }
        void loadingcomplete2()
        {
            form3.Close();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("확인 버튼을 누르시면 아래 표의 내용이 데이터베이스에 저장됩니다!", "경고", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if(dr == DialogResult.OK) 
            {
                form2 = new Form2();
                worker = null;
                worker = new BackgroundWorker();
                worker.WorkerSupportsCancellation = true;
                worker.DoWork += (sender1, args) => updateDB();
                worker.RunWorkerCompleted += (sender1, args) => loadingcomplete();
                worker.RunWorkerAsync();
                form2.ShowDialog();
          
            MessageBox.Show("업데이트 완료!");
            }
        }
        void deleteDB()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Server=150.50.78.152;database=tempdb;uid=sa;pwd=2@@51111";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            conn.Open();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Cells[0].Value.ToString() != "")
                {
                    string companynum = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    string name = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    string carclasscode = dataGridView2.Rows[i].Cells[2].Value.ToString();
                    string carbasecode = dataGridView2.Rows[i].Cells[3].Value.ToString();
                    string changeDate = dataGridView2.Rows[i].Cells[4].Value.ToString();
                    string RegDate = dataGridView2.Rows[i].Cells[6].Value.ToString();
                    string id = dataGridView2.Rows[i].Cells[7].Value.ToString();
                    cmd.CommandText = "IF EXISTS(SELECT * FROM [handicall].[dbo].[DriverInfoScheduleChange] WHERE no = " + companynum + " AND name='" + name + "' AND carclasscode= " + carclasscode + " AND changeDate = '" + changeDate + "') BEGIN DELETE FROM [handicall].[dbo].[DriverInfoScheduleChange] WHERE no = " + companynum + " AND name='" + name + "' AND carclasscode= " + carclasscode + " AND changeDate = '" + changeDate + "' END";
                    cmd.ExecuteNonQuery();
                }
            }
            conn.Close();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("전체 삭제하시겠습니까?", "경고", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                form3 = new Form3();
                worker2 = null;
                worker2 = new BackgroundWorker();
                worker2.WorkerSupportsCancellation = true;
                worker2.DoWork += (sender1, args) => deleteDB();
                worker2.RunWorkerCompleted += (sender1, args) => loadingcomplete2();
                worker2.RunWorkerAsync();
                form3.ShowDialog();
                MessageBox.Show("삭제 완료!");
            }
        }
    }
}
