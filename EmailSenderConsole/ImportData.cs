using ICSharpCode.SharpZipLib.Zip;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EmailSenderConsole
{
    public class ImportData
    {
        //发件人账号
       public List<Sender> Senders = new List<Sender>();

        //收件人账号
       public Stack EmailAccount = new Stack();
       public bool ImportSender(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    string fileType = filePath.Substring(filePath.LastIndexOf(".") + 1);//取得文件后缀
                    FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);//创建文件流
                    bool isXls = true;//判断文件类型
                    if (fileType == "xlsx")
                    {
                        isXls = false;
                    }
                    IWorkbook workbook = CreateWorkbook(isXls, fs);//创建工作簿
                    ISheet sheet = workbook.GetSheetAt(0);//取得第一个工作表
                    int rowCount = sheet.LastRowNum + 1;//取得行数
                    int colCount = sheet.GetRow(0).LastCellNum;//取得列数

                    //DataTable dt = new DataTable();
                    //dt.Columns.Add("Email", Type.GetType("System.String"));
                    //dt.Columns.Add("Key", Type.GetType("System.String"));
                    //ds.Tables.Add(dt);
                    //dataGridView1.DataSource = dt;
                    int r;
                    for (r = 0; r < rowCount; r++)//遍历Excel行,生成dataGridView1单元格内容
                    {//遍历Excel行,从第一行开始
                        IRow row = sheet.GetRow(r);
                        colCount = row.LastCellNum;
                        //DataRow dr = dt.NewRow();
                        //dt.Rows.Add(dr);
                        //dr["Email"] = cell1.ToString();
                        //dr["Key"] = cell2.ToString();
                        ICell cell1 = row.GetCell(0);
                        ICell cell2 = row.GetCell(1);
                        //导入全局
                        Senders.Add(new Sender()
                        {
                            MyEmail = cell1.ToString().Trim(),
                            MyPwd = cell2.ToString().Trim()
                        });
                    }
                    Console.WriteLine("发件账户导入：" + r.ToString());
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("导入失败: " + ex.Message);
                    return false;
                }
            }
            else
            {
                Console.WriteLine("导入失败！");
                return false;
            }
        }
        public bool ImportReceiver(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    string fileType = filePath.Substring(filePath.LastIndexOf(".") + 1);//取得文件后缀
                    FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);//创建文件流
                    bool isXls = true;//判断文件类型
                    if (fileType == "xlsx")
                    {
                        isXls = false;
                    }
                    IWorkbook workbook = CreateWorkbook(isXls, fs);//创建工作簿
                    ISheet sheet = workbook.GetSheetAt(0);//取得第一个工作表
                    int rowCount = sheet.LastRowNum + 1;//取得行数
                    int colCount = sheet.GetRow(0).LastCellNum;//取得列数

                    int r = 0;
                    for (r = 0; r < rowCount; r++)//遍历Excel行,生成dataGridView1单元格内容
                    {//遍历Excel行,从第一行开始
                        IRow row = sheet.GetRow(r);
                        //int index = dataGridView1.Rows.Add();
                        colCount = row.LastCellNum;
                        ICell cell1 = row.GetCell(0);

                        //添加收件人账户
                        EmailAccount.Push(cell1.ToString());
                    }
                    //收件人总数
                    Console.WriteLine("收件人总数：" + r.ToString());

                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("导入失败: " + ex.Message);
                    return false;
                }
            }
            else
            {
                Console.WriteLine("导入失败！");
                return false;
            }
        }
        //}
       private static IWorkbook CreateWorkbook(bool isXLS, FileStream fs)
        {

            if (isXLS)
            {
                return new HSSFWorkbook(fs);
            }
            else
            {
                return new XSSFWorkbook(fs);
            }
        }
    }
}
