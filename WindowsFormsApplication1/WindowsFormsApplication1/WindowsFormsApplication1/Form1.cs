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
using OfficeOpenXml;

namespace WindowsFormsApplication1
{
    /// <summary>
    /// Excle 文件按模板格式转换
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox2.Text = @"C:\Users\gangbeng\Desktop\需替换的图片2\Excel格式化\需要格式化文件";
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
            //获取模板数据信息
            DataSet ds = ToDataTable(textBox1.Text);
            DataTable dt = ds.Tables[0];
            List<String> listColumName = new List<string>();
            if (dt.Columns.Count > 0)
            {
                foreach (DataColumn item in dt.Columns)
                {
                    listColumName.Add(item.ColumnName);
                }
            }
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
            GetFileListByFolder(textBox2.Text, ref listFilePath);

            foreach (String filepath in listFilePath)
            {
                DataSet dsf = ToDataTable(filepath);
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
                            txtMessage.AppendText(filepath.Substring(filepath.LastIndexOf("\\") + 1) + "没有以下列：" + strErrorMessage.Trim('、') + "\r\n");
                            continue;
                        }

                        DataTable newdt = GetDataTableByTempTable(dt, dtf, dic);

                        //新建一个Excle
                        //http://www.cnblogs.com/lwme/archive/2011/11/27/2265323.html
                        //插入数据


                        string stra = filepath.Substring(0, filepath.LastIndexOf("\\"));
                        string strb = "格式化后_" + filepath.Substring(filepath.LastIndexOf("\\") + 1, filepath.LastIndexOf(".") - filepath.LastIndexOf("\\") - 1) + ".xlsx";

                        FileInfo newFile = new FileInfo(stra + "\\" + strb);

                        if (newFile.Exists)
                        {
                            newFile.Delete();  // ensures we create a new workbook
                            newFile = new FileInfo(stra + "\\" + strb);
                        }
                        FileInfo tempfile = new FileInfo(textBox1.Text);
                        using (ExcelPackage package = new ExcelPackage(newFile, tempfile))
                        {
                            var ws = package.Workbook.Worksheets[1];

                            ws.Cells["A1"].LoadFromDataTable(newdt, true);
                            ws.Cells["N2:N" + (dtf.Rows.Count - 1).ToString()].FormulaR1C1 = "RC[-3]*RC[-2]";
                            package.Save();
                            return;
                        }

                        continue;//只取一个Sheet

                    }
                }
            }

        }
        #endregion

        private DataTable GetDataTableByTempTable(DataTable templateTable, DataTable newDataTable, Dictionary<String, String> dic)
        {
            DataTable dt = templateTable.Copy();
            dt.Clear();
            

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




        private void DumpExcel(DataTable tbl)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                //Create the worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Demo");

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(tbl, true);

                ////Format the header for column 1-3
                //using (ExcelRange rng = ws.Cells["A1:C1"])
                //{
                //    rng.Style.Font.Bold = true;
                //    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                //    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                //    rng.Style.Font.Color.SetColor(Color.White);
                //}

                //Example how to Format Column 1 as numeric 
                //using (ExcelRange col = ws.Cells[2, 1, 2 + tbl.Rows.Count, 1])
                //{
                //    col.Style.Numberformat.Format = "#,##0.00";
                //    col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}

                ////Write it back to the client
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("content-disposition", "attachment;  filename=ExcelDemo.xlsx");
                //Response.BinaryWrite(pck.GetAsByteArray());
            }
        }

        #region 遍历文件夹下所有Excle文件
        /// <summary>
        /// 文件夹下
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="listFilePath"></param>
        public static void GetFileListByFolder(String folder, ref List<String> listFilePath)
        {
            DirectoryInfo TheFolder = new DirectoryInfo(folder);
            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {

                string extension = Path.GetExtension(NextFile.FullName);
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (str.Contains(extension))
                {
                    listFilePath.Add(NextFile.FullName);
                }


            }
            //遍历文件夹下
            foreach (DirectoryInfo NextFolder in TheFolder.GetDirectories())
            {
                GetFileListByFolder(NextFolder.FullName, ref listFilePath);
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
                MessageBox.Show("打开模板文件异常", "提示", MessageBoxButtons.OK);
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
    }
}
