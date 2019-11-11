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
using System.Net.Mail;
using System.Net;
using System.Threading;

namespace TBS_bot
{

    public partial class MainWindow : Window
    {
        bool isBotWorking;

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

        string path;
        string pdfFromDocPath;

        List<FlatDescription> flatObjects = new List<FlatDescription>();
        List<FlatDescription> currentFlatObjects = new List<FlatDescription>();

        public MainWindow()
        {
            InitializeComponent();
            isBotWorking = false;
            NotificationTB.Text = "nie działam";
            path = AppDomain.CurrentDomain.BaseDirectory;
            pdfFromDocPath = path + "wniosek.pdf";

            GetEmailSettings();
            ReadJson();
            if (currentFlatObjects != null)
                WriteAdressesList();
        }

        private void ReadJson()
        {

            string jsonPath = path + "json.txt";

            try
            {
                using (StreamReader sr = new StreamReader(jsonPath))
                {
                    string line = sr.ReadToEnd();
                    currentFlatObjects = JsonConvert.DeserializeObject<List<FlatDescription>>(line);
                }
            }
            catch (IOException e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private async void StartBot_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                isBotWorking = !isBotWorking;

                if (isBotWorking)
                {
                    StartBotButton.Content = "Stop Bot";
                    NotificationTB.Text = "działam";
                    MailSettingsGrid.IsEnabled = false;
                }
                else
                {
                    StartBotButton.Content = "Start Bot";
                    NotificationTB.Text = "nie działam";
                    MailSettingsGrid.IsEnabled = true;
                }
            });

            while (isBotWorking)
            {
                await BotTasks();
                await Task.Delay(61000);
            }

        }

        private async Task BotTasks()
        {
            string notiffication = NotificationTB.Text;
            Dispatcher.Invoke(() => { NotificationTB.Text = "Pobieram"; });
            mainWindow.Cursor = Cursors.Wait;


            await GetFlatsList();

            if (IsFlatObjectsDifferent())
            {
                foreach (var item in flatObjects)
                {
                    await GetFlatDescription(item);
                }
                UpdateFlatObjects();

                await SendEmailCheck();

                FlatObjectsSerialize();
                ReadJson();

                AddressesTB.Items.Clear();
                WriteAdressesList();
            }

            Dispatcher.Invoke(() => { NotificationTB.Text = notiffication; });
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
            string result = await GetPageString("http://www.tbs-wroclaw.com.pl/mieszkania-na-wynajem/");
            result = result.Substring(result.IndexOf("<a href=\"http://www.tbs-wroclaw.com.pl/"));
            result = result.Substring(0, result.IndexOf("<p>&nbsp;</p>"));

            string[] paragraphs = result.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in paragraphs)
            {
                string line = item.Replace("<strong>", "");
                line = line.Replace("</strong>", "");
                string temporaryItem = line.Substring(line.IndexOf('"') + 1);
                string link = temporaryItem.Substring(0, temporaryItem.IndexOf('"'));

                temporaryItem = line.Substring(line.IndexOf("\">") + 2);
                string address = temporaryItem.Substring(0, temporaryItem.IndexOf("</"));

                flatObjects.Add(new FlatDescription(address, link));
            }

        }

        private void WriteAdressesList()
        {
            AddressesTB.SelectedItems.Clear();
            AddressesTB.Items.Clear();
            FlatDescriptionTB.Text = "";
            HyperLinkTB.Text = "";

            foreach (var item in currentFlatObjects)
            {
                AddressesTB.Items.Add(item.Address + "\t" + item.IsSend);
            }
        }

        private async Task GetFlatDescription(FlatDescription flat)
        {

            string result = await GetPageString(flat.Link);
            result = result.Substring(result.IndexOf("OGŁOSZENIE"));
            string flatNumber = result.Substring(0, result.IndexOf("</p>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                flatNumber = flatNumber.Replace(item, "");
            }

            flatNumber = flatNumber.Substring(14);

            string address = result.Substring(result.IndexOf("przy ul.") + 8);
            address = address.Replace("&nbsp;", " ");
            address = address.Substring(0, address.IndexOf("we Wrocławiu"));

            result = result.Substring(result.IndexOf("(osiedle") + 9);
            string district = result.Substring(0, result.IndexOf(")"));

            result = result.Substring(result.IndexOf("o powierzchni") + 14);
            double flatArea = Convert.ToDouble(Regex.Replace(result.Substring(0, result.IndexOf("składający się") - 2), "[^0-9,]", ""));

            result = result.Substring(result.IndexOf("<ol>") + 4);
            result = result.Substring(0, result.IndexOf("</ol>"));

            foreach (var item in flatDescriptionReplacementList)
            {
                result = result.Replace(item, "");
            }

            int roomsCount = Regex.Matches(result, "Pokoju").Count;
            int isAneksInt = Regex.Matches(result.ToString(), "Pokoju z aneksem kuchennym").Count;

            bool isAneks;

            if (isAneksInt == 1)
                isAneks = true;

            else
                isAneks = false;

            flat.FlatDescriptionUpdate(flatNumber, address, roomsCount, flatArea, isAneks, district, result);
            SetFlatClassified(flat);
        }

        private void AddressesTB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (AddressesTB.SelectedIndex > -1)
            {
                FlatDescriptionTB.Text = currentFlatObjects.ElementAt(AddressesTB.SelectedIndex).GetDetailedDescription();
                Hyperlink hyperlink = new Hyperlink();
                hyperlink.Inlines.Add(currentFlatObjects.ElementAt(AddressesTB.SelectedIndex).Link);
                hyperlink.Click += new RoutedEventHandler(HyperLinkClick);
                HyperLinkTB.Text = "";
                HyperLinkTB.Inlines.Add(hyperlink);
            }
        }


        private void HyperLinkClick(object sender, RoutedEventArgs e)
        {
            Process.Start(HyperLinkTB.Text);
        }

        private void FlatObjectsSerialize()
        {
            string flatObjectString = JsonConvert.SerializeObject(flatObjects);

            string docPath =
              AppDomain.CurrentDomain.BaseDirectory;

            using (StreamWriter outputFile = new StreamWriter(System.IO.Path.Combine(docPath, "json.txt")))
            {
                outputFile.WriteLine(flatObjectString);
            }
        }

        private void CreateProposal(FlatDescription flat)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            string xlmPath = path + "dane.xlsx";
            string docxPath = path + "wniosek.docx";
            string docxPath_a = path + "wniosek_a.docx";

            xlWorkBook = xlApp.Workbooks.Open(xlmPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = flat.Number;
            xlWorkSheet.Cells[2, 1] = flat.AlternationAddress;
            xlWorkSheet.Cells[3, 1] = flat.RoomsCount;
            xlWorkSheet.Cells[4, 1] = flat.FlatArea;

            xlWorkBook.Saved = true;
            xlWorkBook.SaveCopyAs(xlmPath);

            xlWorkBook.Close(null, null, null);
            xlApp.Quit();

            word.Application app = new word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            word.Document doc = new word.Document();

            if (flat.IsAneks)
                doc = app.Documents.Open(docxPath_a);
            else
                doc = app.Documents.Open(docxPath);

            doc.SaveAs2(pdfFromDocPath, word.WdSaveFormat.wdFormatPDF);
            doc.Close();
            app.Quit();

        }

        private bool IsFlatObjectsDifferent()
        {
            if (currentFlatObjects == null)
                return true;
            else if (flatObjects.Count() != currentFlatObjects.Count())
                return true;
            else
            {
                for (int i = 0; i < flatObjects.Count(); i++)
                {
                    if (!flatObjects.ElementAt(i).Link.Equals(currentFlatObjects.ElementAt(i).Link))
                        return true;
                }
            }
            return false;
        }

        private void UpdateFlatObjects()
        {
            if (currentFlatObjects != null)
                foreach (var item in flatObjects)
                    foreach (var item2 in currentFlatObjects)
                        if (item.Link.Equals(item2.Link))
                        {
                            item.IsSend = item2.IsSend;
                            break;
                        }
        }

        private async Task SendEmailCheck()
        {
            foreach (var item in flatObjects)
            {
                if (!item.IsSend && item.IsClassified)
                {
                    await CreateAndSend(item);
                }
            }
        }

        private async void EmailSender_Click(object sender, RoutedEventArgs e)
        {
            await SendEmail(null);
        }

        private async Task CreateAndSend(FlatDescription item)
        {
            mainWindow.Cursor = Cursors.Wait;
            string notiffication = NotificationTB.Text;
            Dispatcher.Invoke(() => { NotificationTB.Text = "Tworzę wniosek"; });

            while (File.Exists(pdfFromDocPath))
                Thread.Sleep(500);
            CreateProposal(item);
            while (!File.Exists(pdfFromDocPath))
                Thread.Sleep(500);
            while (IsFileLocked(pdfFromDocPath))
                Thread.Sleep(500);

            Dispatcher.Invoke(() => { NotificationTB.Text = "Wysyłam wniosek"; });
            item.IsSend = await SendEmail(item);

            NotificationTB.Text = notiffication;
            mainWindow.Cursor = Cursors.Arrow;
        }

        private async Task<bool> SendEmail(FlatDescription flat)
        {
            string subject = "wiadomość testowa";
            string body = "W załączeniu przesyłam wniosek dot. ogłoszenia nr: ";
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string zalPath = path + "zalTBS.pdf";
            string pdfPath = zalPath;

            if (flat != null)
            {
                subject = flat.Number;
                body += subject;
                pdfPath = pdfFromDocPath;
            }

            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtp = new SmtpClient(ServerSmtpTB.Text);

                mail.From = new MailAddress(EmailTB.Text);
                mail.To.Add(RecievEmailTB.Text);
                mail.Subject = subject;
                mail.Body = body;

                Attachment attachment = new Attachment(pdfPath);
                mail.Attachments.Add(attachment);

                smtp.Port = Convert.ToInt32(SmtpPortTB.Text);
                smtp.Credentials = new NetworkCredential(EmailTB.Text, PasswordTB.Password);
                smtp.EnableSsl = true;

                await smtp.SendMailAsync(mail);

                attachment.Dispose();
                mail.Dispose();

                if (flat == null)
                    MessageBox.Show("Wysłano wiadomość");

                return true;

            }
            catch (Exception ex)
            {
                if (flat == null)
                    MessageBox.Show(ex.ToString());
                else
                    Dispatcher.Invoke(() => { NotificationTB.Text = "błąd email"; });

                return false;
            }
            finally
            {
                if (File.Exists(pdfFromDocPath))
                    File.Delete(pdfFromDocPath);
            }
        }


        private void GetEmailSettings()
        {
            EmailTB.Text = Properties.Settings.Default.email;
            PasswordTB.Password = Properties.Settings.Default.password;
            RecievEmailTB.Text = Properties.Settings.Default.recieverEmail;
            ServerSmtpTB.Text = Properties.Settings.Default.smtpServer;
            SmtpPortTB.Text = Properties.Settings.Default.smtpPort.ToString();
        }

        private void EmailSettingsSaveButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.email = EmailTB.Text;
            Properties.Settings.Default.password = PasswordTB.Password;
            Properties.Settings.Default.recieverEmail = RecievEmailTB.Text;
            Properties.Settings.Default.smtpServer = ServerSmtpTB.Text;
            Properties.Settings.Default.smtpPort = Convert.ToInt32(SmtpPortTB.Text);
            Properties.Settings.Default.Save();
        }

        private bool IsFileLocked(string filePath)
        {
            FileStream fileStream = null;

            try
            {
                FileInfo file = new FileInfo(filePath);
                fileStream = file.Open(FileMode.Open, FileAccess.Write, FileShare.None);

            }
            catch (IOException)
            {

                return true;
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Close();
            }
            return false;
        }

        private void SetFlatClassified(FlatDescription flat)
        {
            if (flat.RoomsCount > 1)
            {
                if (KitchenCB.IsChecked == true && !flat.IsAneks)
                    DistrictClassified(flat);
                else if (KitchenetteCB.IsChecked == true && flat.IsAneks)
                    DistrictClassified(flat);
                else
                    flat.IsClassified = false;
            }
            else
                flat.IsClassified = false;
        }

        private void DistrictClassified(FlatDescription flat)
        {
            switch (flat.District.ToLower())
            {
                case ("stabłowice"):
                    if (StablowicaCB.IsChecked == true)
                        flat.IsClassified = true;
                    break;
                case ("leśnica"):
                    if (LesnicaCB.IsChecked == true)
                        flat.IsClassified = true;
                    break;
                case ("brochów"):
                    if (BrochowCB.IsChecked == true)
                        flat.IsClassified = true;
                    break;
                default:
                    if (OtherCB.IsChecked == true)
                        flat.IsClassified = true;
                    break;
            }
        }

        private async void SendProposalButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddressesTB.SelectedIndex < 0)
            {
                MessageBox.Show("Wybierz mieszkanie do wysłania wniosku");
            }
            else
            {
               await CreateAndSend(currentFlatObjects.ElementAt(AddressesTB.SelectedIndex));
                WriteAdressesList();
            }

        }
    }
}
