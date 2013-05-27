using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace WindowsFormsApplication1
{
    /// <summary>
    /// Excle 文件按模板格式转换
    /// </summary>
    public partial class Form1 : Form
    {

        public static String CurrentPath
        {
            get
            {
                return Directory.GetCurrentDirectory().Replace(@"bin\Debug", "");
            }
        }

        public Form1()
        {
            InitializeComponent();
            txtfileprefix.Text = ConfigurationManager.AppSettings["formatFilePrefix"];
            textBox1.Text = ConfigurationManager.AppSettings["tempFilePath"];
            textBox2.Text = ConfigurationManager.AppSettings["formatFilePath"];
        }

        #region 浏览选文件
        private void btnLookfile_Click(object sender, EventArgs e)
        {
            //初始化一个OpenFileDialog类
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.DefaultExt = "xlsx";
            fileDialog.Filter =
                "Excle files (*.xlsx)|*.xlsx";
            //fileDialog.DefaultExt = ".xls,xlsx";
            fileDialog.Title = "请选择模版文件";



            //判断用户是否正确的选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                //声明允许的后缀名
                string[] str = new string[] { ".xlsx" };
                if (!str.Contains(extension))
                {
                    MessageBox.Show("请选择excle格式文件！", "提示", MessageBoxButtons.OK);
                }
                else
                {
                    //获取用户选择的文件，并判断文件大小不能超过20K，fileInfo.Length是以字节为单位的
                    FileInfo fileInfo = new FileInfo(fileDialog.FileName);
                    textBox1.Text = fileDialog.FileName;
                }
            }
        }
        #endregion

        #region 浏览选文件夹
        private void btnLookfile2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            //folderBrowserDialog.f

            folderBrowserDialog.Description = "选择要格式化的的Excel所存放的文件夹";

            // Do not allow the user to create new files via the FolderBrowserDialog.
            folderBrowserDialog.ShowNewFolderButton = false;

            // Default to the My Documents folder.
            folderBrowserDialog.RootFolder = Environment.SpecialFolder.DesktopDirectory;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog.SelectedPath;
            }


        }
        #endregion

        #region 格式化
        private void btnFormat_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("请选择模板", "提示", MessageBoxButtons.OK);
                return;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("需要格式文件存放文件夹", "提示", MessageBoxButtons.OK);
                return;
            }
            txtMessage.Text = "";


            #region 保存填写信息

#if DEBUG
            string applicationName =
                Environment.GetCommandLineArgs()[0];
#else
           string applicationName =
          Environment.GetCommandLineArgs()[0]+ ".exe";
#endif
            string exePath = System.IO.Path.Combine(
         Environment.CurrentDirectory, applicationName);
            Configuration config = ConfigurationManager.OpenExeConfiguration(exePath);


            //make changes
            config.AppSettings.Settings["formatFilePrefix"].Value = txtfileprefix.Text;
            config.AppSettings.Settings["tempFilePath"].Value = textBox1.Text;
            config.AppSettings.Settings["formatFilePath"].Value = textBox2.Text;

            //save to apply changes
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            #endregion

            string strfilePrefix = String.IsNullOrEmpty(txtfileprefix.Text) ? "格式化后_" : txtfileprefix.Text;

            //获取模板数据信息
            DataSet ds = ToDataTable(textBox1.Text);
            DataTable dt = ds.Tables[0];
            List<String> listColumName = new List<string>();
            //列名
            if (dt.Columns.Count > 0)
            {
                foreach (DataColumn item in dt.Columns)
                {
                    listColumName.Add(item.ColumnName);
                }
            }
            //列名 *……&￥￥
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (!String.IsNullOrEmpty(dr[i].ToString()))
                    {
                        listColumName[i] = listColumName[i] + "," + dr[i].ToString();
                    }
                }
            }


            //遍历文件夹下所有Excel文件
            List<String> listFilePath = new List<string>();
            GetFileListByFolder(textBox2.Text, ref listFilePath, strfilePrefix);
            txtMessage.AppendText("    共有" + listFilePath.Count + "个文件需要格式化...\r\n");

            int icount = 0;
            foreach (String filepath in listFilePath)
            {
                icount++;
                DataSet dsf = ToDataTable(filepath);
                String strShortFilePath = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                txtMessage.AppendText("\r\n正格式化第" + icount + "个文件  " + strShortFilePath + "  ...\r\n");

               
                foreach (DataTable dtf in dsf.Tables)
                {
                    if (dtf.Rows.Count > 1 && dtf.Columns.Count > 1)
                    {
                        //待格式化Excle列名
                        List<String> listColumName2 = new List<string>();
                        if (dtf.Columns.Count > 0)
                        {
                            foreach (DataColumn item in dtf.Columns)
                            {
                                listColumName2.Add(item.ColumnName);
                            }
                        }
                        //列名对应关系
                        Dictionary<String, String> dic = new Dictionary<string, string>();
                        Boolean bCheckOK = true;
                        String strErrorMessage = String.Empty;
                        foreach (string citem in listColumName)
                        {
                            foreach (string newcitem in listColumName2)
                            {
                                if (citem.Replace("*", "").Split(',', '，').Contains(newcitem))
                                {
                                    dic.Add(citem.Split(',', '，')[0], newcitem);
                                    break;
                                }
                            }
                            if (citem.Contains('*') && !dic.ContainsKey(citem.Split(',', '，')[0]))
                            {
                                strErrorMessage += citem.Split(',', '，')[0] + "、";
                                bCheckOK = false;
                            }
                        }
                        if (!bCheckOK)
                        {
                            txtMessage.AppendText("没有以下列：" + strErrorMessage.Trim('、') + "\r\n");
                            continue;
                        }

                        DataTable newdt = GetDataTableByTempTable(dt, dtf, dic);
                        //新建一个Excle
                        string stra = filepath.Substring(0, filepath.LastIndexOf("\\"));
                        string strb = strfilePrefix + filepath.Substring(filepath.LastIndexOf("\\") + 1, filepath.LastIndexOf(".") - filepath.LastIndexOf("\\") - 1) + ".xlsx";
                        FileInfo newFile = new FileInfo(stra + "\\" + strb);
                        if (newFile.Exists)
                        {
                            newFile.Delete();  // ensures we create a new workbook
                            newFile = new FileInfo(stra + "\\" + strb);
                        }
                        FileInfo tempfile = new FileInfo(textBox1.Text);


                        //插入数据
                        using (ExcelPackage package = new ExcelPackage(newFile, tempfile))
                        {
                            var ws = package.Workbook.Worksheets[1];

                            ws.Cells["A1"].LoadFromDataTable(newdt, true);
                            ws.Cells["N2:N" + (dtf.Rows.Count).ToString()].FormulaR1C1 = "RC[-3]*RC[-2]";
                            ws.Cells["M2:M" + (dtf.Rows.Count).ToString()].FormulaR1C1 = "RC[-1]*RC[-3]";
                            ws.Cells[dtf.Rows.Count + 1, 14].Formula = "Sum(N2:N" + (dtf.Rows.Count).ToString() + ")";
                            ws.Cells[dtf.Rows.Count + 1, 13].Formula = "Sum(M2:M" + (dtf.Rows.Count).ToString() + ")";
                            ws.Cells["P2:P" + (dtf.Rows.Count).ToString()].Value = strShortFilePath.Replace(".xls", "").Replace(".xlsx", "");
                            //ws.Cells["N" + (dtf.Rows.Count).ToString()+1].FormulaR1C1 = "SUM(RC[-(dtf.Rows.Count - 1).ToString()],RC[-1]";
                            package.Save();
                            txtMessage.AppendText("√格式化成功" + "\r\n");
                        }


                        //break;//只取一个Sheet

                    }
                }
            }

        }
        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateTable"></param>
        /// <param name="newDataTable"></param>
        /// <param name="dic"></param>
        /// <returns></returns>
        private DataTable GetDataTableByTempTable(DataTable templateTable, DataTable newDataTable, Dictionary<String, String> dic)
        {



            //DataTable dt = templateTable.Copy();
            //dt.Clear();
            DataTable dt = new DataTable();
            DataColumn column;

            foreach (DataColumn item in templateTable.Columns)
            {
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = item.ColumnName;
                dt.Columns.Add(column);
            }


            foreach (var dc in dic)
            {
                if (newDataTable.Rows[0][dc.Value].GetType() != typeof(DBNull))
                {
                    dt.Columns[dc.Key].DataType = newDataTable.Rows[0][dc.Value].GetType();
                }
            }

            foreach (DataRow item in newDataTable.Rows)
            {
                DataRow dr = dt.NewRow();
                foreach (var dc in dic)
                {
                    dr[dc.Key] = item[dc.Value];
                }
                dt.Rows.Add(dr);

            }


            return dt;

        }


        #region 遍历文件夹下所有Excle文件
        /// <summary>
        /// 文件夹下
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="listFilePath"></param>
        public static void GetFileListByFolder(String folder, ref List<String> listFilePath, String strExceptFilePrefix)
        {
            DirectoryInfo TheFolder = new DirectoryInfo(folder);
            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {

                string extension = Path.GetExtension(NextFile.FullName);
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (str.Contains(extension) && !NextFile.Name.Contains(strExceptFilePrefix))
                {
                    listFilePath.Add(NextFile.FullName);
                }


            }
            //遍历文件夹下
            foreach (DirectoryInfo NextFolder in TheFolder.GetDirectories())
            {
                GetFileListByFolder(NextFolder.FullName, ref listFilePath,strExceptFilePrefix);
            }

        }
        #endregion

        #region 读取Excel文件到DataSet中
        /// <summary>  
        /// 读取Excel文件到DataSet中  
        /// </summary>  
        /// <param name="filePath">文件路径</param>  
        /// <returns></returns>  
        public static DataSet ToDataTable(string filePath)
        {
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType)) return null;

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;

            DataSet ds = new DataSet();
            try
            {
                // 初始化连接，并打开  
                conn = new OleDbConnection(connStr);
                conn.Open();

                // 获取数据源的表定义元数据                         
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                // 初始化适配器  
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];

                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }

                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);
                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, SheetName);

                    ds.Tables.Add(dsItem.Tables[0].Copy());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("打开{0}异常 {1}",filePath,ex.ToString()), "提示", MessageBoxButtons.OK);
            }
            finally
            {
                // 关闭连接  
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return ds;
        }
        #endregion

        private void btnshowColumnName_Click(object sender, EventArgs e)
        {

            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("请选择模板", "提示", MessageBoxButtons.OK);
                return;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("需要格式文件存放文件夹", "提示", MessageBoxButtons.OK);
                return;
            }
            txtMessage.Text = "";

            DataSet ds = ToDataTable(textBox1.Text);
            DataTable dt = ds.Tables[0];
            List<String> listColumName = new List<string>();
            //列名
            if (dt.Columns.Count > 0)
            {
                foreach (DataColumn item in dt.Columns)
                {
                    listColumName.Add(item.ColumnName);
                }
            }
            //列名 *……&￥￥
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (!String.IsNullOrEmpty(dr[i].ToString()))
                    {
                        listColumName[i] = listColumName[i] + "," + dr[i].ToString();
                    }
                }
            }
            string strfilePrefix = String.IsNullOrEmpty(txtfileprefix.Text) ? "格式化后_" : txtfileprefix.Text;
            //遍历文件夹下所有Excel文件
            List<String> listFilePath = new List<string>();
            GetFileListByFolder(textBox2.Text, ref listFilePath, strfilePrefix);
            txtMessage.AppendText("    共有" + listFilePath.Count + "个文件需要格式化...\r\n");

            int icount = 0;
            foreach (String filepath in listFilePath)
            {
                icount++;
                DataSet dsf = ToDataTable(filepath);
                String strShortFilePath = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                txtMessage.AppendText("\r\n第" + icount + "个文件  " + strShortFilePath + "  ...\r\n");

               
                foreach (DataTable dtf in dsf.Tables)
                {
                    if (dtf.Rows.Count > 1 && dtf.Columns.Count > 1)
                    {
                        //待格式化Excle列名
                        List<String> listColumName2 = new List<string>();
                        if (dtf.Columns.Count > 0)
                        {
                            foreach (DataColumn item in dtf.Columns)
                            {
                                listColumName2.Add(item.ColumnName);
                            }
                        }
                        txtMessage.AppendText(String.Format("所有列名：\r\n{0}", string.Join("、", listColumName2.ToArray())));
                        //列名对应关系
                        Dictionary<String, String> dic = new Dictionary<string, string>();
                        Boolean bCheckOK = true;
                        String strErrorMessage = String.Empty;
                        foreach (string citem in listColumName)
                        {
                            foreach (string newcitem in listColumName2)
                            {
                                if (citem.Replace("*", "").Split(',', '，').Contains(newcitem))
                                {
                                    dic.Add(citem.Split(',', '，')[0], newcitem);
                                    break;
                                }
                            }
                            if (citem.Contains('*') && !dic.ContainsKey(citem.Split(',', '，')[0]))
                            {
                                strErrorMessage += citem.Split(',', '，')[0] + "、";
                                bCheckOK = false;
                            }
                        }
                        if (!bCheckOK)
                        {
                            txtMessage.AppendText("没有以下列：" + strErrorMessage.Trim('、') + "\r\n");
                            continue;
                        }
                        

                       

                    }
                }
            }
        }
    }
}
