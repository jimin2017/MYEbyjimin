using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using HS.ExcelExt;
using Hs.Tools;
using System.Text.RegularExpressions;


namespace CleanPass
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.CleanPassword();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string url = @"http://www.matrix67.com/blog/feed";
          //  string txt = MyGrab.GetContent(url);
            // System.Windows.Forms.MessageBox.Show(txt);
            Excel.Application app = Globals.ThisAddIn.Application;
            // app.ActiveSheet.range("a1").value = txt;
            app.GetContent(url);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //测试
            Excel.Application app = Globals.ThisAddIn.Application;
       
            string sr = "";
            sr = System.Convert.ToString(System.Windows.Forms.Clipboard.GetText());
           
            int n = 0;
            object[,] arr = new object[100, 11];
            var with_1 = new Regex("\\|.*?\\|");
            MatchCollection col1 = with_1.Matches(sr);
            foreach (Match mm1 in col1)
            {
                arr[n, 0] = mm1.Value;
                n =n+ 1;
            }
            app.Range["a1"].get_Resize(n, 1).Value2 = arr;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            ////数组写到EXCEL 
            //Excel.Application app = Globals.ThisAddIn.Application;
            //int[,] arr = new int[9, 9];
            //for (int i = 0; i < 9; i++)
            //{
            //  for(int j = 0;j < 9;j++)
            //    {
            //        arr[i, j] = (i+1)*(j+1);
            //    }

            //}
            // app.Range["a1" ].get_Resize(9, 9).Value2 = arr;
            Excel.Application app = Globals.ThisAddIn.Application;
            int[] arr = new int[9];
            for (int i = 0; i < 9; i++)
            { arr[i] = i+1; }
            app.Range["a1"].get_Resize(1, 9).Value2 = arr;
            app.Range["a10"].get_Resize(9, 1).Value2=app.WorksheetFunction.Transpose(arr);
        }
    }
}
