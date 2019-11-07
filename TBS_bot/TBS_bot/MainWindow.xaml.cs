using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;

namespace TBS_bot
{

    public partial class MainWindow : Window
    {

        List<string> links = new List<string>();
        List<string> addresses = new List<string>();
        List<string> FlatDescription = new List<string>();
        List<string> flatDescriptionReplacementList = new List<string>{
            "<li>",
            "</li>",
            "<sup>",
            "</sup>",
            " &nbsp;",
            "&nbsp;",
            "\t",
             "<strong>",
              "</strong>",
              "<u>",
              "</u>",
              "<p>",
              "</p>",
        };

        List<FlatDescription> flatObjects = new List<FlatDescription>();
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void GetFlatsList_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.Cursor = Cursors.Wait;
            addressesTB.Items.Clear();

            string Result = await GetPageString("http://www.tbs-wroclaw.com.pl/mieszkania-na-wynajem/");
            GetFlatsList(Result);

            foreach (var item in links)
            {
                Result = await GetPageString(item);
                FlatDescription.Add(getFlatDescription(Result));
            }
            mainWindow.Cursor = Cursors.Arrow;
        }

        private async Task<string> GetPageString(string url)
        {
            HttpClient httpClient = new HttpClient();
            string Result = await httpClient.GetStringAsync(url);
            return Result;
        }

        private void GetFlatsList(string Result)
        {
            Result = Result.Substring(Result.IndexOf("<p><strong><a href=\""));
            Result = Result.Substring(0, Result.IndexOf("<p>&nbsp;</p>"));
            string[] Paragraphs = Result.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in Paragraphs)
            {
                string line = item.Replace("<strong>", "");
                line = line.Replace("</strong>", "");
                string TemporaryItem = line.Substring(line.IndexOf('"') + 1);
                TemporaryItem = TemporaryItem.Substring(0, TemporaryItem.IndexOf('"'));
                links.Add(TemporaryItem);
                TemporaryItem = line.Substring(line.IndexOf("\">") + 2);
                TemporaryItem = TemporaryItem.Substring(0, TemporaryItem.IndexOf("</"));
                addresses.Add(TemporaryItem);
            }

            foreach (var item in addresses)
            {
                addressesTB.Items.Add(item);
            }
        }

        private string getFlatDescription(string Result)
        {
            StringBuilder stringBuilder = new StringBuilder();

            Result = Result.Substring(Result.IndexOf("OGŁOSZENIE"));
            string flatNumber = Result.Substring(0, Result.IndexOf("</p>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                flatNumber = flatNumber.Replace(item, "");
            }
            flatNumber = flatNumber.Substring(14);
            stringBuilder.Append("ogłoszenie nr: " + flatNumber + Environment.NewLine);

            string address = Result.Substring(Result.IndexOf("przy ul.") + 8);
            address = address.Replace("&nbsp;", " ");
            address = address.Substring(0, address.IndexOf("we Wrocławiu"));
            stringBuilder.Append(address + Environment.NewLine);

            Result = Result.Substring(Result.IndexOf("(osiedle") + 1);
            stringBuilder.Append(Result.Substring(0, Result.IndexOf(")")) + Environment.NewLine);

            Result = Result.Substring(Result.IndexOf("o powierzchni") + 14);

            double flatArea = Convert.ToDouble(Regex.Replace(Result.Substring(0, Result.IndexOf("składający się") - 2), "[^0-9,]", ""));

            stringBuilder.Append("pow: " + flatArea + Environment.NewLine);

            Result = Result.Substring(Result.IndexOf("<ol>") + 4);
            Result = Result.Substring(0, Result.IndexOf("</ol>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                Result = Result.Replace(item, "");
            }

            stringBuilder.Append(Result + Environment.NewLine);
            stringBuilder.Append("part: " + (flatArea * 1200).ToString("F") + Environment.NewLine);
            stringBuilder.Append("czynsz: " + (flatArea * 14.25).ToString("F") + Environment.NewLine);

            int RoomsCount = Regex.Matches(stringBuilder.ToString(), "Pokoju").Count;
            int isAneksInt = Regex.Matches(stringBuilder.ToString(), "Pokoju z aneksem kuchennym").Count;

            stringBuilder.Append("ilośc pokoi: " + RoomsCount + Environment.NewLine);
            bool isAneks;

            if (isAneksInt == 1) {
                stringBuilder.Append("z aneksem: true" + Environment.NewLine);
                isAneks = true;
            }
            else {
                stringBuilder.Append("z aneksem: false" + Environment.NewLine);
                isAneks = false;
            }

            FlatDescription fd = new FlatDescription(flatNumber,address,RoomsCount,flatArea,isAneks,false);
            flatObjects.Add(fd);

            return stringBuilder.ToString();
        }

        private void AddressesTB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            linksTB.Text = FlatDescription.ElementAt(addressesTB.SelectedIndex);
            Hyperlink hyperlink = new Hyperlink();
            hyperlink.Inlines.Add(links.ElementAt(addressesTB.SelectedIndex));
            hyperlink.Click += new RoutedEventHandler(HyperLinkClick);
            hyperLinkTB.Text = "";
            hyperLinkTB.Inlines.Add(hyperlink);
        }

        private void HyperLinkClick(object sender, RoutedEventArgs e)
        {
            Process.Start(hyperLinkTB.Text);
        }

        private void FlatObjectsSerialize()
        {
            string test = JsonSerializer.Serialize(flatObjects);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            string path = AppDomain.CurrentDomain.BaseDirectory;
            string xlmPath =  path + "dane.xlsx";
            string docxPath = path + "wniosek.docx";
            string pdfPath = path + "wniosek.pdf";
            xlWorkBook = xlApp.Workbooks.Open(xlmPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //xlWorkSheet.Cells[1, 1] = "test";

            xlWorkSheet.Cells[1, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).number;
            xlWorkSheet.Cells[2, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).Addres;
            xlWorkSheet.Cells[3, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).flatNymbers;
            xlWorkSheet.Cells[4, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).flatArea;

            xlWorkBook.Saved = true;
            xlWorkBook.SaveCopyAs(xlmPath);

            xlWorkBook.Close(null,null,null);
            xlApp.Quit();

            word.Application app = new word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            word.Document doc = app.Documents.Open(docxPath);
     
            doc.SaveAs2(pdfPath, word.WdSaveFormat.wdFormatPDF);
            doc.Close();
            app.Quit();

            //releaseObject(xlWorkSheet);
           // releaseObject(xlWorkBook);
           // releaseObject(xlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
