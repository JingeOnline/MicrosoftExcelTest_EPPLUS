using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using OfficeOpenXml;

namespace MicrosoftExcelTest
{
    /// <summary>
    /// EPPLUS单元格的index和Excel一样，都是从1开始的。
    /// </summary>
    class Program
    {
        static string filePath = @"C:\Users\jinge\Desktop\test1.xlsx";

        static void Main(string[] args)
        {
            //必须加上该行，否则在debug模式下，会报LicenseException异常。
            //其他的License验证方式见 https://github.com/EPPlusSoftware/EPPlus
            //当前使用的验证方式是用的添加appSetting.json文件的方式
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //create();
            //read();
            //update();
            //insertRow();
            //deleteRow();
            //loadCsv();
            loadObject();
        }

        //创建Excel文件并写入数据
        static void create()
        {
            using (var package = new ExcelPackage())
            {
                //添加一个Sheet
                var worksheet = package.Workbook.Worksheets.Add("Inventory");
                //添加表头（使用行列数定位单元格）
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Value";

                //添加每三行的数据（使用单元格坐标定位单元格）
                worksheet.Cells["A2"].Value = 12001;
                worksheet.Cells["B2"].Value = "Nails";
                worksheet.Cells["C2"].Value = 37;
                worksheet.Cells["D2"].Value = 3.99;

                worksheet.Cells["A3"].Value = 12002;
                worksheet.Cells["B3"].Value = "Hammer";
                worksheet.Cells["C3"].Value = 5;
                worksheet.Cells["D3"].Value = 12.10;

                worksheet.Cells["A4"].Value = 12003;
                worksheet.Cells["B4"].Value = "Saw";
                worksheet.Cells["C4"].Value = 12;
                worksheet.Cells["D4"].Value = 15.37;

                //为Value列添加公式
                worksheet.Cells["E2:E4"].Formula = "C2*D2";
                //执行计算（一般不需要，打开Excel时会自动计算。但如果不用Excel打开，则通过预先计算，就能直接储存计算结果。）
                worksheet.Calculate();
                //自动设置列宽（可传入最小和最大宽度）
                worksheet.Cells.AutoFitColumns();
                //设置文件属性
                package.Workbook.Properties.Title = "Invertory";
                package.Workbook.Properties.Author = "Jan Källman";
                package.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel workbook using EPPlus";
                package.Workbook.Properties.Company = "EPPlus Software AB";

                //保存并输出文件
                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }
            Console.WriteLine("Create file finish.");
        }

        //读取Excel文件数据
        static void read()
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //打印当前表格的行数
                Console.WriteLine("Rows=" + worksheet.Dimension.Rows);
                //打印当前表格范围
                Console.WriteLine("Dimension=" + worksheet.Dimension.Address);
                //打印表格最后一行和最后一列的index
                Console.WriteLine("Last cell row, " + worksheet.Dimension.End.Row);
                Console.WriteLine("Last cell column, " + worksheet.Dimension.End.Column);
                //循环打印2-4行的第二列单元格
                int col = 2; //Column 2 is the item description
                for (int row = 2; row < 5; row++)
                {
                    Console.WriteLine("\tCell({0},{1}), Value={2}", row, col, worksheet.Cells[row, col].Value);
                }
                //打印单元格中的公式（普通格式和R1C1格式）
                Console.WriteLine("\tCell({0},{1}).Formula={2}", 3, 5, worksheet.Cells[3, 5].Formula);
                Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells[3, 5].FormulaR1C1);
            }
            Console.WriteLine("Read file finish.");
        }

        //修改单元格中的数据
        static void update()
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //修改一个单元格
                worksheet.Cells[2, 4].Value = 100;
                //修改一行数据(从该行第二列开始修改)
                worksheet.Cells[3, 2].LoadFromText("Spanner,25,25");
                //保存
                package.SaveAs(file);
            }
            Console.WriteLine("Update file finish.");
        }

        //插入新行（下面的行会自动下移）
        static void insertRow()
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //从第二行开始，插入两个空行，原先存在的行会自动下移
                worksheet.InsertRow(2, 2);
                package.SaveAs(file);
            }
            Console.WriteLine("Insert row finish.");
        }

        //删除行（下面的行会自动上移）
        static void deleteRow()
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //从第二行开始，插入两个空行，原先存在的行会自动下移
                worksheet.DeleteRow(2, 2);
                package.SaveAs(file);
            }
            Console.WriteLine("Delete row finish.");
        }

        //从CSV导入数据到表
        static void loadCsv()
        {
            string csvText = "Id,Name,Quantity,Price,Value\n10001,sss,12,12,32\n10002,lll,14,14,14";
            //设置csv解析格式
            ExcelTextFormat format = new ExcelTextFormat
            {
                //换行符
                EOL = "\n",
                //字段分隔符
                Delimiter = ',', 
                //解析时，跳过的行数
                SkipLinesBeginning = 1,
            };
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //找到当前文件的最后一行
                int endRow = worksheet.Dimension.End.Row;
                //从新的一行开始写入数据
                worksheet.Cells[endRow + 1, 1].LoadFromText(csvText,format);
                package.Save();
            }
            Console.WriteLine("Load CSV finish.");
        }

        //从对象列表导入数据到表
        static void loadObject()
        {
            List<MyData> list = new List<MyData>
            {
                new MyData{col1="0001",col2="windows",col3=10,col4=10,col5=10.1 },
                new MyData{col1="0002",col2="ios",col3=1,col4=1,col5=11.1 },
                new MyData{col1="0002",col2="android",col3=5,col4=5,col5=15.1 },
            };
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //读取文件中的第1个sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //找到当前文件的最后一行
                int endRow = worksheet.Dimension.End.Row;
                //从新的一行开始写入数据
                worksheet.Cells[endRow + 1, 1].LoadFromCollection(list);
                package.Save();
            }
            Console.WriteLine("Load from object list finish");
        }
    }



    public class MyData
    {
        public string col1 { get; set; }
        public string col2 { get; set; }
        public int col3 { get; set; }
        public int col4 { get; set; }
        public double col5 { get; set; }
    }
}
