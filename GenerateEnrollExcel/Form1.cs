using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Data.OleDb;
using System.Threading;

namespace GenerateEnrollExcel
{
    public partial class Frm_Generator : Form
    {
        public Frm_Generator()
        {
            InitializeComponent();
        }

        private void btn_select_Click(object sender, EventArgs e)
        {
            var result = ofD_file.ShowDialog();
            txtFile.Text = ofD_file.FileName;
        }

        private void btn_generator_Click(object sender, EventArgs e)
        {
            var source = ExcelToDS(txtFile.Text);
            var classNames = source.Select(o => new { o.className , o.school }).Distinct().ToList();

            foreach (var curClass in classNames)
            {
                XSSFWorkbook hssfworkbook = null;
                using (FileStream fs = File.Open(@"template.xlsx", FileMode.Open,
                FileAccess.Read, FileShare.ReadWrite))
                {
                    //把xls文件读入workbook变量里，之后就可以关闭了  
                    hssfworkbook = new XSSFWorkbook(fs);
                    fs.Close();
                }

                var curSource = source.Where(o => o.className == curClass.className && o.school==curClass.school).ToList();
                var fileName = string.Format("{0}_{1}_{2}_{3}.xlsx", curSource[0].school, curSource[0].grade, curSource[0].subject, curClass.className);

                XSSFSheet sheet1 = hssfworkbook.GetSheet("template") as XSSFSheet;
                hssfworkbook.SetSheetName(0, curClass.className);
                sheet1.GetRow(2).GetCell(0).SetCellValue("课程名称： " + curClass.className);
                sheet1.GetRow(2).GetCell(6).SetCellValue("上课时间： " + curSource[0].courseTime);

                //var curSource = source.Where(o => o.className == curClass.className && o.school==curClass.school).ToList();
                for (var i = 0; i < curSource.Count; i++)
                {
                    //sheet1.GetRow(i + 5).GetCell(0).SetCellValue(i+1);
                    sheet1.GetRow(i + 5).GetCell(1).SetCellValue(curSource[i].studentName);
                    sheet1.GetRow(i + 5).GetCell(5).SetCellValue(curSource[i].mobile);
                    sheet1.GetRow(i + 5).GetCell(6).SetCellValue(curSource[i].studentSchool);
                    sheet1.GetRow(i + 5).GetCell(7).SetCellValue(curSource[i].studentClass);
                }
                
                sheet1.ForceFormulaRecalculation = true;

                using (FileStream fileStream = File.Open(fileName,
                    FileMode.Create, FileAccess.ReadWrite))
                {
                    hssfworkbook.Write(fileStream);
                    fileStream.Close();
                }
            }
            MessageBox.Show("导出成功！");
        }

        public List<Order> ExcelToDS(string Path)
        {
            FileStream file = new FileStream(Path, FileMode.Open, FileAccess.Read);
            XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
            XSSFSheet sheet1 = hssfworkbook.GetSheet("报名情况") as XSSFSheet;
            List<Order> orders = new List<Order>();
            var i = 1;
            while(sheet1.GetRow(i)!=null)
            {
                var newOrder = new Order();
                for (var j = 0; j < 11; j++)
                {
                    newOrder.studentName = sheet1.GetRow(i).GetCell(0).StringCellValue;
                    newOrder.mobile = sheet1.GetRow(i).GetCell(1).StringCellValue;
                    newOrder.className = sheet1.GetRow(i).GetCell(9).StringCellValue;
                    newOrder.school = sheet1.GetRow(i).GetCell(8).StringCellValue;
                    newOrder.grade = sheet1.GetRow(i).GetCell(11).StringCellValue;
                    newOrder.subject = sheet1.GetRow(i).GetCell(7).StringCellValue;
                    newOrder.courseTime = sheet1.GetRow(i).GetCell(10).StringCellValue;
                    newOrder.studentSchool = (sheet1.GetRow(i).GetCell(2)!=null?sheet1.GetRow(i).GetCell(2).StringCellValue:"");
                    newOrder.studentClass = (sheet1.GetRow(i).GetCell(3)!=null?sheet1.GetRow(i).GetCell(3).StringCellValue:"");
                }
                if(newOrder.className!="")
                {
                    orders.Add(newOrder);
                }                
                i++;
            }

             return orders;
        }

        private void btnAddMobile_Click(object sender, EventArgs e)
        {
            FileStream file = new FileStream(txtFile.Text, FileMode.Open, FileAccess.ReadWrite);
            XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
            XSSFSheet sheet1 = hssfworkbook.GetSheet("老师") as XSSFSheet;
            List<Teacher> teachers = new List<Teacher>();
            var i = 1;
            while (sheet1.GetRow(i) != null)
            {
                var teacher = new Teacher();
                teacher.name = sheet1.GetRow(i).GetCell(0).StringCellValue;
                teacher.mobile = sheet1.GetRow(i).GetCell(1).NumericCellValue;

                if (teacher.name != "")
                {
                    teachers.Add(teacher);
                }
                i++;
            }

            XSSFSheet sheet2 = hssfworkbook.GetSheet("Sheet1") as XSSFSheet;

            i = 1;
            while (sheet2.GetRow(i) != null)
            {
                var mobile  = sheet2.GetRow(i).GetCell(6)!=null?sheet2.GetRow(i).GetCell(6).NumericCellValue:0;
                var name = sheet2.GetRow(i).GetCell(5)!=null? sheet2.GetRow(i).GetCell(5).StringCellValue:"";

                if (name != "" && mobile == 0)
                {
                    var findTeachers = teachers.FindAll(o => o.name == name);
                    if (findTeachers.Count == 1)
                    {
                        mobile = findTeachers[0].mobile;
                        sheet2.GetRow(i).GetCell(6).SetCellValue(mobile);
                    }
                    else if (findTeachers.Count > 1)
                    {
                        sheet2.GetRow(i).GetCell(6).SetCellValue("重名了");
                    }
                    else
                    {
                        sheet2.GetRow(i).GetCell(6).SetCellValue("没找到");
                    }
                }
                i++;
            }
            file = new FileStream(txtFile.Text, FileMode.Open, FileAccess.ReadWrite);
            hssfworkbook.Write(file);
            file.Flush();
            file.Close();
        }
    }

    public class Order
    {
        public string studentName;
        public string mobile;
        public string studentSchool;
        public string studentClass;
        public string sex;
        public string subject;
        public string school;
        public string className;
        public string grade;
        public string courseTime;
    }

    public class Teacher
    {
        public string name;
        public double mobile;
    }

    //    var source = ExcelToDS(txtFile.Text);
    //    var classNames = source.Select(o => new { o.className, o.school }).Distinct().ToList();

    //    FileStream file = new FileStream(@"template.xlsx", FileMode.Open, FileAccess.Read);
    //    XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
    //    XSSFSheet sheetBase = hssfworkbook.GetSheet("template") as XSSFSheet;
    //    hssfworkbook.RemoveSheetAt(0);

    //            //var count = 0;
    //            foreach (var curClass in classNames)
    //            {
    //                var curSource = source.Where(o => o.className == curClass.className && o.school == curClass.school).ToList();
    //    var fileName = string.Format("{0}_{1}_{2}_{3}.xlsx", curSource[0].school, curSource[0].grade, curSource[0].subject, curClass.className);

    //    //File.Copy("template.xlsx", fileName, true);

    //    //FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
    //    //XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
    //    //XSSFSheet sheet1 = hssfworkbook.GetSheet("template") as XSSFSheet;
    //    var sheet1 = sheetBase.CopySheet(curClass.className);
    //    hssfworkbook.Add(sheet1);
    //                //hssfworkbook.SetSheetName(0, curClass.className);
    //                sheet1.GetRow(2).GetCell(0).SetCellValue("课程名称： " + curClass.className);
    //    sheet1.GetRow(2).GetCell(6).SetCellValue("上课时间： " + curSource[0].courseTime);

    //                //var curSource = source.Where(o => o.className == curClass.className && o.school==curClass.school).ToList();
    //                for (var i = 0; i<curSource.Count; i++)
    //                {
    //                    //sheet1.GetRow(i + 5).GetCell(0).SetCellValue(i+1);
    //                    sheet1.GetRow(i + 5).GetCell(1).SetCellValue(curSource[i].studentName);
    //    sheet1.GetRow(i + 5).GetCell(5).SetCellValue(curSource[i].mobile);
    //}

    //sheet1.ForceFormulaRecalculation = true;

    //                FileStream fileOut = new FileStream(fileName, FileMode.Create);
    //hssfworkbook.Write(fileOut);
    //                fileOut.Close();

    //                hssfworkbook.RemoveSheetAt(0);
    //                //if(count==5)
    //                //{
    //                //    break;
    //                //}
    //                //count++;
    //            }
    //            MessageBox.Show("导出成功！");
}
