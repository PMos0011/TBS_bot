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
using Excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;
using System.IO;
using Newtonsoft.Json;

namespace TBS_bot
{

    public partial class MainWindow : Window
    {

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
        List<FlatDescription> CurrentFlatObjects = new List<FlatDescription>();

        public MainWindow()
        {
            InitializeComponent();
            ReadJson();
            if(CurrentFlatObjects!=null)
                WriteAdressesList();
        }

        private void ReadJson()
        {
            string docPath =
              AppDomain.CurrentDomain.BaseDirectory;
            docPath += "json.txt";

            try
            {
                using (StreamReader sr = new StreamReader(docPath))
                {
                    string line = sr.ReadToEnd();
                    CurrentFlatObjects = JsonConvert.DeserializeObject<List<FlatDescription>>(line);
                }
            }
            catch (IOException e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private async void GetFlatsList_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.Cursor = Cursors.Wait;
            
            ReadJson();
 
            await GetFlatsList();

            if (isFlatObjectsDifferent())
            {
                foreach (var item in flatObjects)
                {
                    await GetFlatDescription(item);
                }
                UpdateFlatObjects();

                SendEmail();

                FlatObjectsSerialize();
                ReadJson();

                addressesTB.Items.Clear();
                WriteAdressesList();
            }

            mainWindow.Cursor = Cursors.Arrow;

        }

        private async Task<string> GetPageString(string url)
        {
            HttpClient httpClient = new HttpClient();
            string Result = await httpClient.GetStringAsync(url);
            return Result;
        }

        private async Task GetFlatsList()
        {
            flatObjects.Clear();
            string Result = await GetPageString("http://www.tbs-wroclaw.com.pl/mieszkania-na-wynajem/");
            Result = Result.Substring(Result.IndexOf("<a href=\"http://www.tbs-wroclaw.com.pl/"));
            Result = Result.Substring(0, Result.IndexOf("<p>&nbsp;</p>"));

            string[] Paragraphs = Result.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in Paragraphs)
            {
                string line = item.Replace("<strong>", "");
                line = line.Replace("</strong>", "");
                string TemporaryItem = line.Substring(line.IndexOf('"') + 1);
                string link = TemporaryItem.Substring(0, TemporaryItem.IndexOf('"'));

                TemporaryItem = line.Substring(line.IndexOf("\">") + 2);
                string address = TemporaryItem.Substring(0, TemporaryItem.IndexOf("</"));

                flatObjects.Add(new FlatDescription(address, link));
            }

        }

        private void WriteAdressesList()
        {
            addressesTB.SelectedItems.Clear();
            addressesTB.Items.Clear();
            linksTB.Text = "";
            hyperLinkTB.Text = "";

            foreach (var item in CurrentFlatObjects)
            {
                addressesTB.Items.Add(item.Address + "\t" + item.isSend);
            }
        }

        private async Task GetFlatDescription(FlatDescription flatDescrition)
        {

            string Result = await GetPageString(flatDescrition.Link);
            Result = Result.Substring(Result.IndexOf("OGŁOSZENIE"));
            string flatNumber = Result.Substring(0, Result.IndexOf("</p>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                flatNumber = flatNumber.Replace(item, "");
            }

            flatNumber = flatNumber.Substring(14);

            string address = Result.Substring(Result.IndexOf("przy ul.") + 8);
            address = address.Replace("&nbsp;", " ");
            address = address.Substring(0, address.IndexOf("we Wrocławiu"));

            Result = Result.Substring(Result.IndexOf("(osiedle") + 1);
            string district = Result.Substring(0, Result.IndexOf(")")) + Environment.NewLine;

            Result = Result.Substring(Result.IndexOf("o powierzchni") + 14);
            double flatArea = Convert.ToDouble(Regex.Replace(Result.Substring(0, Result.IndexOf("składający się") - 2), "[^0-9,]", ""));

            Result = Result.Substring(Result.IndexOf("<ol>") + 4);
            Result = Result.Substring(0, Result.IndexOf("</ol>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                Result = Result.Replace(item, "");
            }

            int RoomsCount = Regex.Matches(Result, "Pokoju").Count;
            int isAneksInt = Regex.Matches(Result.ToString(), "Pokoju z aneksem kuchennym").Count;

            bool isAneks;

            if (isAneksInt == 1)
                isAneks = true;

            else
                isAneks = false;

            flatDescrition.FlatDescriptionUpdate(flatNumber, address, RoomsCount, flatArea, isAneks, false, district + Result);
        }

        private void AddressesTB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (addressesTB.SelectedIndex > -1)
            {
                linksTB.Text = CurrentFlatObjects.ElementAt(addressesTB.SelectedIndex).GetDetailedDescription();
                Hyperlink hyperlink = new Hyperlink();
                hyperlink.Inlines.Add(CurrentFlatObjects.ElementAt(addressesTB.SelectedIndex).Link);
                hyperlink.Click += new RoutedEventHandler(HyperLinkClick);
                hyperLinkTB.Text = "";
                hyperLinkTB.Inlines.Add(hyperlink);
            }
        }
        

        private void HyperLinkClick(object sender, RoutedEventArgs e)
        {
            Process.Start(hyperLinkTB.Text);
        }

        private void FlatObjectsSerialize()
        {
            string FlatObjectString = JsonConvert.SerializeObject(flatObjects);

            string docPath =
              AppDomain.CurrentDomain.BaseDirectory;

            using (StreamWriter outputFile = new StreamWriter(System.IO.Path.Combine(docPath, "json.txt")))
            {
                outputFile.WriteLine(FlatObjectString);
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            string path = AppDomain.CurrentDomain.BaseDirectory;
            string xlmPath = path + "dane.xlsx";
            string docxPath = path + "wniosek.docx";
            string docxPath_a = path + "wniosek_a.docx";
            string pdfPath = path + "wniosek.pdf";
            xlWorkBook = xlApp.Workbooks.Open(xlmPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).number;
            xlWorkSheet.Cells[2, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).AlternationAddress;
            xlWorkSheet.Cells[3, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).RoomsCount;
            xlWorkSheet.Cells[4, 1] = flatObjects.ElementAt(addressesTB.SelectedIndex).flatArea;

            xlWorkBook.Saved = true;
            xlWorkBook.SaveCopyAs(xlmPath);

            xlWorkBook.Close(null, null, null);
            xlApp.Quit();

            word.Application app = new word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            word.Document doc = new word.Document();

            if (flatObjects.ElementAt(addressesTB.SelectedIndex).isAneks)
                doc = app.Documents.Open(docxPath_a);
            else
                doc = app.Documents.Open(docxPath);

            doc.SaveAs2(pdfPath, word.WdSaveFormat.wdFormatPDF);
            doc.Close();
            app.Quit();

        }

        private bool isFlatObjectsDifferent()
        {
            if (CurrentFlatObjects == null)
                return true;
            else if (flatObjects.Count() != CurrentFlatObjects.Count())
                return true;
            else
            {
                for (int i = 0; i < flatObjects.Count(); i++)
                {
                    if (!flatObjects.ElementAt(i).Link.Equals(CurrentFlatObjects.ElementAt(i).Link))
                        return true;
                }
            }
            return false;
        }

        private void UpdateFlatObjects()
        {
            if (CurrentFlatObjects != null)
                foreach (var item in flatObjects)
                    foreach (var item2 in CurrentFlatObjects)
                        if (item.Link.Equals(item2.Link))
                        {
                            item.isSend = item2.isSend;
                            break;
                        }
        }

        private void SendEmail()
        {
            foreach (var item in flatObjects)
            {
                if (!item.isSend)
                    if (item.RoomsCount > 1)
                        item.isSend = true;
            }
        }

        private void EmailSender_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
