using org.in2bits.MyXls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 易之星车贷数据处理
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        #region private member
        /// <summary>  
        /// 新建文件保存路径  
        /// </summary>  
        private string savePath = string.Empty;

        /// <summary>  
        /// 报账数据表数据集  
        /// </summary>  
        private DataSet oneXLSDs = null;

        /// <summary>  
        /// 业务总表表数据集  
        /// </summary>  
        private DataSet twoXLSDs = null;

        /// <summary>  
        /// 支付明细表数据集  
        /// </summary>  
        private DataSet threeXLSDs = null;

        /// <summary>  
        /// 服务费台账表数据集  
        /// </summary>  
        private DataSet fourXLSDs= null;

        /// <summary>  
        ///  数据库连接  
        /// </summary>  
        private OleDbConnection conn;

        /// <summary>  
        /// 源文件  
        /// </summary>  
        private string orinFileName = string.Empty;

        /// <summary>  
        /// 目标文件  
        /// </summary>  
        private string[] aimFileName = new string[5];

        /// <summary>  
        /// 批量读取  
        /// </summary>  
        private string[] everyRead = new string[300];
        #endregion

        #region 读取文件路径

        /// <summary>  
        /// 添加目的文件1  
        /// </summary>  
        /// <param name="sender">消息对象</param>  
        /// <param name="e">消息体</param>  
        private void file_one_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel表格(*.xls)|*.xls";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.ShowDialog();
            string selectFileName = openFileDialog1.FileName;

            // 文件名显示框  
            file_one_name.Text = selectFileName;

            // 保存文件名  
            this.aimFileName[0] = file_one_name.Text.Remove(0, selectFileName.LastIndexOf("\\") + 1);
        }

        /// <summary>  
        /// 添加目的文件2  
        /// </summary>  
        /// <param name="sender">消息对象</param>  
        /// <param name="e">消息体</param>  
        private void file_two_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel表格(*.xls)|*.xls";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.ShowDialog();
            string selectFileName = openFileDialog1.FileName;

            // 文件名显示框  
            file_two_name.Text = selectFileName;

            // 保存文件名  
            this.aimFileName[1] = file_one_name.Text.Remove(0, selectFileName.LastIndexOf("\\") + 1);
        }

        /// <summary>  
        /// 添加目的文件3  
        /// </summary>  
        /// <param name="sender">消息对象</param>  
        /// <param name="e">消息体</param>  
        private void file_three_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel表格(*.xls)|*.xls";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.ShowDialog();
            string selectFileName = openFileDialog1.FileName;

            // 文件名显示框  
            file_three_name.Text = selectFileName;

            // 保存文件名  
            this.aimFileName[2] = file_one_name.Text.Remove(0, selectFileName.LastIndexOf("\\") + 1);
        }

        /// <summary>  
        /// 添加目的文件4  
        /// </summary>  
        /// <param name="sender">消息对象</param>  
        /// <param name="e">消息体</param>  
        private void file_out_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel表格(*.xls)|*.xls";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.ShowDialog();
            string selectFileName = openFileDialog1.FileName;

            // 文件名显示框  
            file_four_name.Text = selectFileName;

            // 保存文件名  
            this.aimFileName[3] = file_one_name.Text.Remove(0, selectFileName.LastIndexOf("\\") + 1);
        }
        #endregion

        #region 开始比对
        /// <summary>  
        /// 比对  
        /// </summary>  
        /// <param name="sender">消息对象</param>  
        /// <param name="e">消息体</param>  
        private void start_check_Click(object sender, EventArgs e)
        {
            try
            {
                //1.解析报账数据表
                this.AnalysisOneXLS(file_one_name.Text);
                //2.解析业务总表
                this.AnalysisTwoXLS(file_two_name.Text);
                //3.解析支行上传表
                this.AnalysisThreeXLS(file_three_name.Text);
                //4.解析校验结果表
                this.AnalysisFourXLS(file_four_name.Text);
                //5.表一是否支付
                this.CheckPay(oneXLSDs, twoXLSDs);
                //6.表五
                this.fillTable(oneXLSDs, twoXLSDs, threeXLSDs);
                //7.表六
                this.CompareTable(oneXLSDs, fourXLSDs);
                //8.删除总表，重新生成  
                this.SaveFile(this.twoXLSDs.Tables[0], file_two_name.Text);
                //9.删除校验结果表，重新生成
                this.SaveFile(this.threeXLSDs.Tables[0], file_three_name.Text);
                //10.删除校验结果表，重新生成
                this.SaveFile(this.fourXLSDs.Tables[0], file_four_name.Text);
                MessageBox.Show("完成");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        #endregion

        #region 解析数据
        /// <summary>  
        /// 解析报账数据表   
        /// </summary>  
        /// <param name="filePath">报账数据表全路径</param>
        private void AnalysisOneXLS(string filePath)
        {
            try
            {
                // 获得连接  
                string connStr = this.GetConnStr(filePath);

                if (System.IO.File.Exists(filePath))
                {
                    OleDbDataAdapter myCommand = null;
                    DataSet ds = null;

                    using (this.conn = new OleDbConnection(connStr))
                    {
                        this.conn.Open();

                        // 得到所有sheet的名字  
                        DataTable sheetsName = this.conn.GetOleDbSchemaTable(
                            OleDbSchemaGuid.Tables,
                            new object[] { null, null, null, "Table" });

                        // 得到第一个sheet的名字  
                        string firstSheetName = sheetsName.Rows[0][2].ToString();
                        string strExcel = "Select * From  [" + firstSheetName + "A:K]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.oneXLSDs = ds;
                    }

                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>  
        /// 解析业务总表  
        /// </summary>  
        /// <param name="filePath">业务总表全路径</param>
        private void AnalysisTwoXLS(string filePath)
        {
            try
            {
                // 获得连接  
                string connStr = this.GetConnStr(filePath);

                if (System.IO.File.Exists(filePath))
                {
                    OleDbDataAdapter myCommand = null;
                    DataSet ds = null;

                    using (this.conn = new OleDbConnection(connStr))
                    {
                        this.conn.Open();

                        // 得到所有sheet的名字  
                        DataTable sheetsName = this.conn.GetOleDbSchemaTable(
                            OleDbSchemaGuid.Tables,
                            new object[] { null, null, null, "Table" });

                        // 得到第一个sheet的名字  
                        string firstSheetName = sheetsName.Rows[0][2].ToString();
                        string strExcel = "Select * From  [" + firstSheetName + "]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.twoXLSDs = ds;
                    }

                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>  
        /// 解析支付明细表
        /// </summary>  
        /// <param name="filePath">支付明细表全路径</param>
        private void AnalysisThreeXLS(string filePath)
        {
            try
            {
                // 获得连接  
                string connStr = this.GetConnStr(filePath);

                if (System.IO.File.Exists(filePath))
                {
                    OleDbDataAdapter myCommand = null;
                    DataSet ds = null;

                    using (this.conn = new OleDbConnection(connStr))
                    {
                        this.conn.Open();

                        // 得到所有sheet的名字  
                        DataTable sheetsName = this.conn.GetOleDbSchemaTable(
                            OleDbSchemaGuid.Tables,
                            new object[] { null, null, null, "Table" });

                        // 得到第一个sheet的名字  
                        string firstSheetName = sheetsName.Rows[0][2].ToString();
                        string strExcel = "Select * From  [" + firstSheetName + "]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.threeXLSDs = ds;
                    }

                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>  
        /// 解析服务费台账表  
        /// </summary>  
        /// <param name="filePath">服务费台账表全路径</param>
        private void AnalysisFourXLS(string filePath)
        {
            try
            {
                // 获得连接  
                string connStr = this.GetConnStr(filePath);

                if (System.IO.File.Exists(filePath))
                {
                    OleDbDataAdapter myCommand = null;
                    DataSet ds = null;

                    using (this.conn = new OleDbConnection(connStr))
                    {
                        this.conn.Open();

                        // 得到所有sheet的名字  
                        DataTable sheetsName = this.conn.GetOleDbSchemaTable(
                            OleDbSchemaGuid.Tables,
                            new object[] { null, null, null, "Table" });

                        // 得到第一个sheet的名字  
                        string firstSheetName = sheetsName.Rows[0][2].ToString();
                        string strExcel = "Select * From  [" + firstSheetName + "]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.fourXLSDs= ds;
                    }

                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        #endregion

        #region 保存文件

        /// <summary>  
        /// 保存文件  
        /// </summary>  
        /// <param name="dt">DataTable</param>  
        /// <param name="fullPath">保存的文件路径</param>  
        private void SaveFile(DataTable dt, string fullPath)
        {
            // 存在该文件，就删除，重新建立  
            if (File.Exists(fullPath))
            {
                File.Delete(fullPath);
            }

            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            CreateExcel(fullPath, dt);
        }
        #endregion

        #region 表一是否支付
        /// <summary>  
        /// 表一是否支付
        /// </summary>  
        /// <param name="oneDs">外循环Ds</param>  
        /// <param name="twoDs">内循环Ds</param>  
        private void CheckPay(DataSet oneDs, DataSet twoDs)
        {
            string data = string.Empty;
            for (int i = 0; i < oneDs.Tables[0].Rows.Count; i++)
            {
                data = string.Empty;
                string cell = oneDs.Tables[0].Rows[i][4].ToString().Trim();
                for (int j = 0; j < twoDs.Tables[0].Rows.Count; j++)
                {
                    if (cell == twoDs.Tables[0].Rows[j][14].ToString().Trim() && oneDs.Tables[0].Rows[i][10].ToString().Trim()=="")
                    {
                        twoDs.Tables[0].Rows[j][21] = "已支付";
                        twoDs.Tables[0].Rows[j][18] = DateTime.Now.ToString("yyyyMMdd");
                    }
                }
            }
        }
        #endregion

        #region 表五登记
        /// <summary>  
        /// 表五登记 
        /// </summary>  
        /// <param name="twoDs,oneDs">外循环Ds</param>  
        /// <param name="threeDs">内循环Ds</param>  
        private void fillTable(DataSet oneDs, DataSet twoDs, DataSet threeDs)
        {
            string data = string.Empty;
            for (int i = 0; i < twoDs.Tables[0].Rows.Count; i++)
            {
                data = string.Empty;
                if (threeDs.Tables[0].Rows.Count <= i)
                {
                    threeDs.Tables[0].Rows.Add();
                }
                for (int j = 0; j < twoDs.Tables[0].Columns.Count - 1; j++)
                {
                    // 数据填充
                    threeDs.Tables[0].Rows[i][j] = twoDs.Tables[0].Rows[i][j];
                }
            }
            for (int i = 0; i < oneDs.Tables[0].Rows.Count; i++)
            {
                string cell = oneDs.Tables[0].Rows[i][4].ToString().Trim();
                for (int j = 0; j < threeDs.Tables[0].Rows.Count; j++)
                {
                    if (cell == threeDs.Tables[0].Rows[j][14].ToString().Trim() && oneDs.Tables[0].Rows[i][8].ToString().Trim() != "" && oneDs.Tables[0].Rows[i][9].ToString().Trim() != "")
                    {
                        threeDs.Tables[0].Rows[j][22] = oneDs.Tables[0].Rows[i][8].ToString().Trim();
                        threeDs.Tables[0].Rows[j][23] = oneDs.Tables[0].Rows[i][9].ToString().Trim();
                    }
                }
            }
        }
        #endregion

        #region 台账生成
        /// <summary>  
        /// 校验比对  
        /// </summary>  
        /// <param name="checkDs">外循环Ds</param>  
        /// <param name="simpleDs">内循环Ds</param>  
        private void CompareTable(DataSet oneDs, DataSet fourDs)
        {
            int k = 0;
            for (int i = oneDs.Tables[0].Rows.Count - 1; i >= 0; i--)
            {
                if (fourDs.Tables[0].Rows.Count <= k)
                {
                    fourDs.Tables[0].Rows.Add();
                }
                if (i >= oneDs.Tables[0].Rows.Count)
                {
                    continue;
                }
                fourDs.Tables[0].Rows[k][1] = oneDs.Tables[0].Rows[i][1].ToString().Trim();
                fourDs.Tables[0].Rows[k][2] = oneDs.Tables[0].Rows[i][2].ToString().Trim();
                fourDs.Tables[0].Rows[k][3] = DateTime.Now.ToString("yyyyMMdd");
                string cell = oneDs.Tables[0].Rows[i][1].ToString().Trim();
                for (int j = oneDs.Tables[0].Rows.Count - 1; j >= 0; j--)
                {
                    if (cell == oneDs.Tables[0].Rows[j][1].ToString().Trim())
                    {
                        fourDs.Tables[0].Rows[k][4] = StrToFloat(fourDs.Tables[0].Rows[k][4].ToString().Trim()) + StrToFloat(oneDs.Tables[0].Rows[j][6].ToString().Trim());
                        fourDs.Tables[0].Rows[k][5] = StrToFloat(fourDs.Tables[0].Rows[k][5].ToString().Trim()) + StrToFloat(oneDs.Tables[0].Rows[j][9].ToString().Trim());
                        oneDs.Tables[0].Rows.Remove(oneDs.Tables[0].Rows[j]);
                    }
                }
                k++;
            }

        }
        #endregion
       
        #region 其他

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("请保证操作的文件没有被打开,并确保各文件首行为列名");
        }

        public int IsNumeric(string str)
        {
            int i;
            if (str != null && System.Text.RegularExpressions.Regex.IsMatch(str, @"^-?\d+$"))
                i = int.Parse(str);
            else
                i = -1;
            return i;
        }
        public float StrToFloat(object FloatString)
        {
            try
            {
                float f = (float)Convert.ToSingle(FloatString);
                return f;
            }
            catch (FormatException)
            {
                return (float)0.00;
            }
        }
        #endregion

        #region 导出函数
        public void CreateExcel(string fileName, DataTable DataSource)
        {
            XlsDocument xls = new XlsDocument();

            #region 列标题样式
            XF columnTitleXF = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            columnTitleXF.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            columnTitleXF.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            columnTitleXF.UseBorder = true; // 使用边框   
            columnTitleXF.TopLineStyle = 1; // 上边框样式  
            columnTitleXF.TopLineColor = Colors.Black; // 上边框颜色  
            columnTitleXF.BottomLineStyle = 1; // 下边框样式  
            columnTitleXF.BottomLineColor = Colors.Black; // 下边框颜色  
            columnTitleXF.LeftLineStyle = 1; // 左边框样式  
            columnTitleXF.LeftLineColor = Colors.Black; // 左边框颜色  
            columnTitleXF.RightLineStyle = 1; // 右边框样式  
            columnTitleXF.RightLineColor = Colors.Black; // 右边框颜色  
            columnTitleXF.Font.Bold = true; // 是否加楚
            #endregion

            #region 数据单元格样式
            XF dataXF = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            dataXF.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            dataXF.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            dataXF.UseBorder = true; // 使用边框   
            dataXF.TopLineStyle = 1; // 上边框样式  
            dataXF.TopLineColor = Colors.Black; // 上边框颜色  
            dataXF.BottomLineStyle = 1; // 下边框样式  
            dataXF.BottomLineColor = Colors.Black; // 下边框颜色  
            dataXF.LeftLineStyle = 1; // 左边框样式  
            dataXF.LeftLineColor = Colors.Black; // 左边框颜色  
            dataXF.RightLineStyle = 1; // 右边框样式  
            dataXF.RightLineColor = Colors.Black; // 右边框颜色  
            dataXF.Font.FontName = "宋体";
            dataXF.Font.Height = 12 * 20; // 设定字大小（字体大小是以 1/20 point 为单位的）  
            dataXF.UseProtection = false; // 默认的就是受保护的，导出后需要启用编辑才可修改  
            dataXF.TextWrapRight = true; // 自动换行  
            #endregion

            Worksheet sheet;
            sheet = xls.Workbook.Worksheets.Add("export_1_");
            ColumnInfo columnInfo = new ColumnInfo(xls, sheet);
            columnInfo.ColumnIndexStart = 0;
            columnInfo.ColumnIndexEnd = (ushort)(DataSource.Columns.Count - 1);
            columnInfo.Width = 15 * 330;
            sheet.AddColumnInfo(columnInfo);

            // 开始填充数据到单元格  
            Cells cells = sheet.Cells;

            for (int j = 0; j < DataSource.Columns.Count; j++)
            {
                cells.Add(1, j + 1, DataSource.Columns[j].ColumnName, columnTitleXF);
            }

            if (DataSource.Rows.Count == 0)
            {
                cells.Add(3, 1, "没有可用数据");
                return;
            }

            for (int l = 0; l < DataSource.Rows.Count; l++)
            {
                for (int j = 0; j < DataSource.Columns.Count; j++)
                {
                    if (j == 0)
                    {
                        cells.Add(l + 2, j + 1, l+1, dataXF);
                    }
                    else
                    {
                        cells.Add(l + 2, j + 1, DataSource.Rows[l][j].ToString(), dataXF);
                    }
                }

            }

            int fileI = fileName.LastIndexOf("\\");
            xls.FileName = fileName.Substring(fileI + 1, fileName.Length - (fileI + 1));
            string path = fileName.Substring(0, fileI + 1);
            try
            {
                xls.Save(path, true);
            }
            catch (Exception)
            {
                throw new MyException("未导出成功，文件可能被占用");
            }
        }
        #endregion

        #region 自定义异常
        public class MyException : ApplicationException
        {
            private string error;
            private Exception innerException;
            //无参数构造函数
            public MyException()
            {

            }
            //带一个字符串参数的构造函数，作用：当程序员用Exception类获取异常信息而非 MyException时把自定义异常信息传递过去
            public MyException(string msg)
                : base(msg)
            {
                this.error = msg;
            }
            //带有一个字符串参数和一个内部异常信息参数的构造函数
            public MyException(string msg, Exception innerException)
                : base(msg)
            {
                this.innerException = innerException;
                this.error = msg;
            }
            public string GetError()
            {
                return error;
            }
        }
        #endregion

        #region 连接语句
        /// <summary>  
        /// 获取连接sql语句  
        /// </summary>  
        /// <param name="filePath">报表全路径</param>  
        /// <returns>连接sql语句</returns>  
        private string GetConnStr(string filePath)
        {
            string connStr = string.Empty;

            if (filePath.Contains(".xls"))
            {
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
            }
            else
            {
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath.Remove(filePath.LastIndexOf("\\") + 1) + ";Extended Properties='Text;FMT=Delimited;HDR=YES;'";
            }

            return connStr;
        }
        #endregion  
    }
}
