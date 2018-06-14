using org.in2bits.MyXls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace 易之星车贷数据处理
{
    /// <summary>  
    ///  窗口类  
    /// </summary>  
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region private member  
        /// <summary>  
        /// 新建文件保存路径  
        /// </summary>  
        private string savePath = string.Empty;  

        /// <summary>  
        /// 存量业务总表数据集  
        /// </summary>  
        private DataSet totalXLSDs = null;  
  
        /// <summary>  
        /// 经销商金融服务费标准表数据集  
        /// </summary>  
        private DataSet severpriceXLSDs = null;

        /// <summary>  
        /// 星之宝报账数据表数据集  
        /// </summary>  
        private DataSet updatecountXLSDs = null;

        /// <summary>  
        /// 查验结果表数据集  
        /// </summary>  
        private DataSet checkedcountXLSDs = null;  
  
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
                //1.解析总表
                this.AnalysisTotalXLS(file_one_name.Text);
                //2.解析服务费率标准表
                this.AnalysisSeverpriceXLS(file_two_name.Text);
                //3.解析支行上传表
                this.AnalysisUpdatecountXLS(file_three_name.Text);
                //4.解析校验结果表
                this.AnalysisCheckedcountXLS(file_four_name.Text);
                // 运行 backgroundWorker 组件
                this.backgroundWorker1.RunWorkerAsync();
                // 显示进度条窗体
                ProcessForm form = new ProcessForm(this.backgroundWorker1, updatecountXLSDs.Tables[0].Rows.Count);
                form.ShowDialog(this);
                form.Close();
                //7.删除总表，重新生成  
                //this.SaveFile(this.totalXLSDs.Tables[0], file_one_name.Text);
                //8.删除校验结果表，重新生成
                this.SaveFile(this.checkedcountXLSDs.Tables[0], file_four_name.Text);
                MessageBox.Show("完成");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
            
        }
        #endregion

        #region 进度条
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
            }
            else
            {
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)//后台工作进度
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            try
            {
                //5.填充校验结果表
                this.fillTable(updatecountXLSDs, checkedcountXLSDs);
                //6.比较总表与校验结果表
                this.CompareTable(checkedcountXLSDs, totalXLSDs,sender,e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 解析数据
        /// <summary>  
        /// 解析总表   
        /// </summary>  
        /// <param name="filePath">支付报表全路径</param>
        private void AnalysisTotalXLS(string filePath)
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
                        //string a="序号,交易检索号,交易地区号,机构名称,特约商户编号,商户名称（交易场所）,特约商户部门编号,商户方手续费,贴息行业,贴息类型,专项分期项目,客户姓名,证件类型,证件号码,交易卡号,交易日期,分期金额,分期期数,最后录入日期,信息查看,信息匹配,是否已支付";
                        string strExcel = "Select * From  [" + firstSheetName + "]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.totalXLSDs = ds;
                    }

                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw(ex);
            }
        }
        /// <summary>  
        /// 解析经销商金融服务费标准  
        /// </summary>  
        /// <param name="filePath">支付报表全路径</param>
        private void AnalysisSeverpriceXLS(string filePath)
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
                        string strExcel = "Select 特约商户编号,调整日期1,服务费标准1,调整日期2,服务费标准2,调整日期3,服务费标准3,调整日期4,服务费标准4,调整日期5,服务费标准5 From  [" + firstSheetName + "]";
                        myCommand = new OleDbDataAdapter(strExcel, this.conn);

                        ds = new DataSet();
                        myCommand.Fill(ds, "Sheet");

                        this.severpriceXLSDs = ds;
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
        /// 解析星之宝报账数据
        /// </summary>  
        /// <param name="filePath">支付报表全路径</param>
        private void AnalysisUpdatecountXLS(string filePath)
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

                        this.updatecountXLSDs = ds;
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
        /// 解析查验结果  
        /// </summary>  
        /// <param name="filePath">支付报表全路径</param>
        private void AnalysisCheckedcountXLS(string filePath)
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

                        this.checkedcountXLSDs = ds;
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

        #region 表四数据填充
        /// <summary>  
        /// 表四数据填充  
        /// </summary>  
        /// <param name="outDs">外循环Ds</param>  
        /// <param name="inDs">内循环Ds</param>  
        private void fillTable(DataSet outDs, DataSet inDs)
        {
            string data = string.Empty; 
            for (int i = 0; i < outDs.Tables[0].Rows.Count; i++)
            {
                data = string.Empty;
                if (inDs.Tables[0].Rows.Count<=i)
                {
                    inDs.Tables[0].Rows.Add();
                }
                for (int j = 0; j < inDs.Tables[0].Columns.Count-3; j++)
                {
                    // 数据填充
                    inDs.Tables[0].Rows[i][j] = outDs.Tables[0].Rows[i][j];
                }
            }
        }
        #endregion  

        #region 校验比对
        /// <summary>  
        /// 校验比对  
        /// </summary>  
        /// <param name="checkDs">外循环Ds</param>  
        /// <param name="simpleDs">内循环Ds</param>  
        private void CompareTable(DataSet checkDs, DataSet simpleDs, Object sender, DoWorkEventArgs e)
        {
            int a = 0;
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int i = 0; i < checkDs.Tables[0].Rows.Count; i++)
            {
                worker.ReportProgress(i);
                if (worker.CancellationPending)  // 如果用户取消则跳出处理数据代码 
                {
                    e.Cancel = true;
                    break;
                }
                string cell = checkDs.Tables[0].Rows[i][4].ToString().Trim();
                bool flag = false;
                for (int j = 0; j < simpleDs.Tables[0].Rows.Count; j++)
                {
                    a = 0;
                    // 比对交易日期，分期金额是否匹配
                    if (cell == simpleDs.Tables[0].Rows[j][14].ToString().Trim())
                    {
                        flag = true;
                        if (simpleDs.Tables[0].Rows[i][21].ToString().Trim() == "已支付")
                        {
                            checkDs.Tables[0].Rows[i][10] += "该笔服务费已发";
                            break;
                        }
                        if (checkDs.Tables[0].Rows[i][5].ToString().Trim() != simpleDs.Tables[0].Rows[j][15].ToString().Trim())
                        {
                            checkDs.Tables[0].Rows[i][10] += "交易日期有误";
                            a++;
                        }
                        if (checkDs.Tables[0].Rows[i][6].ToString().Trim() != simpleDs.Tables[0].Rows[j][16].ToString().Trim())
                        {
                            checkDs.Tables[0].Rows[i][10] += "分期金额有误";
                            a++;
                        }
                        if (checkDs.Tables[0].Rows[i][7].ToString().Trim() != simpleDs.Tables[0].Rows[j][17].ToString().Trim())
                        {
                            checkDs.Tables[0].Rows[i][10] += "分期期数有误";
                            a++;
                        }
                        if (checkDs.Tables[0].Rows[i][1].ToString().Trim() != simpleDs.Tables[0].Rows[j][4].ToString().Trim())
                        {
                            checkDs.Tables[0].Rows[i][10] += "特约商户编号有误";
                            a++;
                        }
                        if (a == 0)
                        {
                            this.fillsevercash(checkDs, severpriceXLSDs, simpleDs, i);
                        }
                        break;
                    }
                }
                if(flag==false)
                {
                    checkDs.Tables[0].Rows[i][10] += "未找到符合的身份证号";
                }
            }
        }
        #endregion  

        #region 计算服务费
        /// <summary>  
        /// 计算服务费  
        /// </summary>  
        /// <param name="sumDs">外循环Ds</param>  
        /// <param name="simpleDs">内循环Ds</param>  
        private void fillsevercash(DataSet sumDs, DataSet simpleDs, DataSet totalDs, int i)
        {
            int type = 0;
            bool flag=false;
            string cell = sumDs.Tables[0].Rows[i][1].ToString().Trim();
            type = 0;
            for (int j = 0; j < simpleDs.Tables[0].Rows.Count; j++)
            {

                if (cell == simpleDs.Tables[0].Rows[j][0].ToString().Trim())
                {
                    flag=true;
                    if (this.IsNumeric(sumDs.Tables[0].Rows[i][5].ToString().Trim()) <= this.IsNumeric(simpleDs.Tables[0].Rows[j][1].ToString().Trim()))
                    {
                        type = 1;
                    }
                    else if (this.IsNumeric(sumDs.Tables[0].Rows[i][5].ToString().Trim()) <= this.IsNumeric(simpleDs.Tables[0].Rows[j][3].ToString().Trim()))
                    {
                        type = 2;
                    }
                    else if (this.IsNumeric(sumDs.Tables[0].Rows[i][5].ToString().Trim()) <= this.IsNumeric(simpleDs.Tables[0].Rows[j][5].ToString().Trim()))
                    {
                        type = 3;
                    }
                    else if (this.IsNumeric(sumDs.Tables[0].Rows[i][5].ToString().Trim()) <= this.IsNumeric(simpleDs.Tables[0].Rows[j][7].ToString().Trim()))
                    {
                        type = 4;
                    }
                    else if (this.IsNumeric(sumDs.Tables[0].Rows[i][5].ToString().Trim()) <= this.IsNumeric(simpleDs.Tables[0].Rows[j][9].ToString().Trim()))
                    {
                        type = 5;
                    }
                    else
                    {
                        sumDs.Tables[0].Rows[i][10] += "交易日期超出服务费调整日期";
                    }
                    switch (type)
                    {

                        case 1:
                            {
                                sumDs.Tables[0].Rows[i][8] = simpleDs.Tables[0].Rows[j][2].ToString().Trim();
                                sumDs.Tables[0].Rows[i][9] = (this.IsNumeric(sumDs.Tables[0].Rows[i][6].ToString().Trim()) * this.StrToFloat(sumDs.Tables[0].Rows[i][8].ToString().Trim())).ToString();
                                break;
                            }
                        case 2:
                            {
                                sumDs.Tables[0].Rows[i][8] = simpleDs.Tables[0].Rows[j][4].ToString().Trim();
                                sumDs.Tables[0].Rows[i][9] = (this.IsNumeric(sumDs.Tables[0].Rows[i][6].ToString().Trim()) * this.StrToFloat(sumDs.Tables[0].Rows[i][8].ToString().Trim())).ToString();
                                break;
                            }
                        case 3:
                            {
                                sumDs.Tables[0].Rows[i][8] = simpleDs.Tables[0].Rows[j][6].ToString().Trim();
                                sumDs.Tables[0].Rows[i][9] = (this.IsNumeric(sumDs.Tables[0].Rows[i][6].ToString().Trim()) * this.StrToFloat(sumDs.Tables[0].Rows[i][8].ToString().Trim())).ToString();
                                break;
                            }
                        case 4:
                            {
                                sumDs.Tables[0].Rows[i][8] = simpleDs.Tables[0].Rows[j][8].ToString().Trim();
                                sumDs.Tables[0].Rows[i][9] = (this.IsNumeric(sumDs.Tables[0].Rows[i][6].ToString().Trim()) * this.StrToFloat(sumDs.Tables[0].Rows[i][8].ToString().Trim())).ToString();
                                break;
                            }
                        case 5:
                            {
                                sumDs.Tables[0].Rows[i][8] = simpleDs.Tables[0].Rows[j][10].ToString().Trim();
                                sumDs.Tables[0].Rows[i][9] = (this.IsNumeric(sumDs.Tables[0].Rows[i][6].ToString().Trim()) * this.StrToFloat(sumDs.Tables[0].Rows[i][8].ToString().Trim())).ToString();
                                break;
                            }
                        default: break;
                    }
                    break;
                }
            }
            if (flag == false)
            {
                sumDs.Tables[0].Rows[i][10] += "未匹配到相应的服务费标准";
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
                cells.Add(1, j+1 , DataSource.Columns[j].ColumnName, columnTitleXF);
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
                    cells.Add(l + 2, j + 1,DataSource.Rows[l][j].ToString(), dataXF);
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
