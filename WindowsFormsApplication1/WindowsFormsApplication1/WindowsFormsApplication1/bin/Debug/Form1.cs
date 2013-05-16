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
        }

        #region 浏览选文件
        private void btnLookfile_Click(object sender, EventArgs e)
        {
            //初始化一个OpenFileDialog类
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.DefaultExt = "xls";
            fileDialog.Filter =
                "Excle files (*.xls)|*.xls|Excle files (*.xlsx)|*.xlsx";
            //fileDialog.DefaultExt = ".xls,xlsx";
            fileDialog.Title = "请选择模版文件";



            //判断用户是否正确的选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (!str.Contains(extension))
                {
                    MessageBox.Show("请选择excle格式文件！","提示",MessageBoxButtons.OK);
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
            folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop; 

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
                   if(!String.IsNullOrEmpty(dr[i].ToString()))
                   {
                       listColumName[i] = listColumName[i] + "," + dr[i].ToString();
                   }
                }
            }

            //遍历文件夹下所有Excel文件
            List<String> listFilePath = new List<string>();
            GetFileListByFolder(textBox2.Text, ref listFilePath);
            Int32 iFileCount = listFilePath.Count;



            
            
        }
        #endregion

        #region 遍历文件夹下所有Excle文件
        /// <summary>
        /// 文件夹下
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="listFilePath"></param>
        public static void GetFileListByFolder(String folder,  ref List<String> listFilePath)
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
