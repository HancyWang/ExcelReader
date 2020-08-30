using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<List<string>> m_list_filtered = new List<List<string>>();
        private List<string> m_list_title = new List<string>();
        private int m_process_val = 0;
        private const string FIND_TAB_STR = "数据源";

        private void button_load_file_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (m_list_filtered.Count > 0 || m_list_title.Count > 0)
                {
                    m_list_filtered.Clear();
                    m_list_title.Clear();
                    button_export_data.Enabled = false;
                    progressBar1.Value = 0;
                    m_process_val = 0;
                }
                

                this.textBox_filePath.Text = this.openFileDialog1.FileName;
                //FileStream stream = new FileStream(this.openFileDialog1.FileName, FileMode.Open);

                using (FileStream stream = new FileStream(this.openFileDialog1.FileName, FileMode.Open))
                {
                    var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    var result = reader.AsDataSet();
                    // MessageBox.Show(result.Tables[0].Rows[0][0].ToString());

                    //获取改excel表中，table的个数
                    int tables_count = result.Tables.Count;
                    string str_tables = "";
                    int tab_index = 0;
                    for (int i = 0; i < tables_count; i++)
                    {
                        str_tables += result.Tables[i].TableName;
                        str_tables += "\n";

                        if (result.Tables[i].TableName == FIND_TAB_STR)
                        {
                            tab_index = i;
                            //MessageBox.Show(tab_index.ToString());
                            break;

                        }
                    }
                   // MessageBox.Show(str_tables);

                    //解析table[1] "数据源"
                    DataTable datetable_src = result.Tables[tab_index];
                    int row_count = datetable_src.Rows.Count;     //行的个数
                    int col_count = datetable_src.Columns.Count;  //列的个数

                    string[] row_element = new string[col_count];

                    //1.获取所有的数据
                    List<List<string>> list_all = new List<List<string>>();
                    for (int i = 0; i < 1; i++)             //获取title
                    {
                        for (int j = 0; j < col_count; j++)
                        {
                            m_list_title.Add((result.Tables[tab_index].Rows[i][j].ToString().Trim())); //记得去除空格
                        }
                    }


                    int INDEX_CONNECT_TYPE = 12;
                    int INDEX_WORKSHOP_NO = 0;
                    int INDEX_PART_NO = 2;
                    int INDEX_NUM = 6;
                    int INDEX_TOTAL_NUM = 7;
                    for (int i = 1; i < row_count; i++)     //获取数据
                    {
                        List<string> list_tmp = new List<string>();
                        for (int j = 0; j < col_count; j++)
                        {
                            string str = result.Tables[tab_index].Rows[i][j].ToString().Trim();
                            //容错
                            if (str == "")
                            {
                                str = "0";
                            }
                            else //如果包含"," 则将","改成";"
                            {
                                str=str.Replace(',', ';');
                            }
                            list_tmp.Add(str); //记得去除空格
                        }
                        list_all.Add(list_tmp);
                    }

                    //2.另起一个链表,去重,相同的就叠加
                    //获取去重后的 接头类型_车间_单号。。。
                    //List<List<string>> m_list_filtered = new List<List<string>>();
                    for (int i = 0; i < list_all.Count; i++)
                    {
                        m_process_val = i;

                        List<string> list_tmp1 = new List<string>();
                        for (int j = 0; j < col_count; j++)
                        {
                            list_tmp1.Add(list_all[i][j]);
                        }
                        if (m_list_filtered.Count > 0)
                        {
                            bool b_new = false;
                            int count = 0;
                            for (int idx=0;idx< m_list_filtered.Count;idx++)
                            {
                                if (list_tmp1[INDEX_CONNECT_TYPE] == m_list_filtered[idx][INDEX_CONNECT_TYPE]
                                        && list_tmp1[INDEX_WORKSHOP_NO] == m_list_filtered[idx][INDEX_WORKSHOP_NO]
                                        && list_tmp1[INDEX_PART_NO] == m_list_filtered[idx][INDEX_PART_NO])
                                {
                                    m_list_filtered[idx][INDEX_NUM] = Convert.ToString(Convert.ToInt32(m_list_filtered[idx][INDEX_NUM]) + Convert.ToInt32(list_tmp1[INDEX_NUM]));
                                    m_list_filtered[idx][INDEX_TOTAL_NUM] = Convert.ToString(Convert.ToInt32(m_list_filtered[idx][INDEX_TOTAL_NUM]) + Convert.ToInt32(list_tmp1[INDEX_TOTAL_NUM]));
                                }
                                else
                                {
                                    count++;
                                }
                                if (count == m_list_filtered.Count)
                                {
                                    b_new = true;
                                }
                            }
                            if (b_new)
                            {
                                m_list_filtered.Add(list_tmp1);
                            }
                        }
                        else
                        {
                            //加入第一个数据
                            m_list_filtered.Add(list_tmp1);
                        }
                        
                    }
                    progressBar1.Maximum = list_all.Count;
                    progressBar1.Minimum = 0;
                    //MessageBox.Show(m_list_filtered.Count.ToString());


                    //打开 "导出数据"
                    button_export_data.Enabled = true;
                }
                //Console.Read();
            }
        }



        public DataSet ExcelToDS(string Path)
        {
            try
            {

                //连接字符串
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串                    //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set);
                    //return set.Tables[0];
                    return set;
                }
            }
            catch (Exception)
            {
                return null;
            }

        }

        private void button_export_data_Click(object sender, EventArgs e)
        {
            //if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var path = this.saveFileDialog1.FileName;
                //path += @"\"+ "导出报表.xls";
                

                FileStream fs = new FileStream(path, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

                string str_title = "";
                for (int i = 0; i < m_list_title.Count; i++)
                {
                    str_title += m_list_title[i] + ",";
                }
                //str_title += "\r\n";
                sw.WriteLine(str_title);

                for (int i = 0; i < m_list_filtered.Count; i++)
                {
                    string str_tmp = "";
                    for (int j = 0; j < m_list_filtered[0].Count; j++)
                    {
                        str_tmp += m_list_filtered[i][j] + ",";
                    }
                    //str_tmp += "\r\n";
                    sw.WriteLine(str_tmp);
                }

                sw.Close();
                fs.Close();
                MessageBox.Show("导出报表:"+path+" 成功。");
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = m_process_val;
        }
    }
}
