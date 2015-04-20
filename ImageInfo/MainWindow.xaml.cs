using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using System.Net;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Drawing;
using System.Drawing.Printing;

namespace ImageInfo
{
    class FileData
    {
        public string Name {get; set; }
        public double Size {get; set; }

        public FileData()
        { Name = ""; Size = 0; }

        public FileData(string name, double size)
        { this.Name = name; this.Size = size; }

        public override string ToString()
        {
            return String.Format("{0}: {1} Mb", Name, Math.Round(Size,5));
        }

    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string imageDir;
        private Dictionary<string, List<FileData>> data = new Dictionary<string, List<FileData>>();
        private double givenSize = 1;

        public MainWindow()
        {
            InitializeComponent();

            //Initialize output
            txtOutput.Text = "";


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Ask the user which directory to scan
            //Depends on Windows API Code Pack
            CommonOpenFileDialog dlg = new CommonOpenFileDialog();
            dlg.Title = "Select a directory with tests";
            dlg.IsFolderPicker = true;
            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                data.Clear();
                imageDir = dlg.FileName;

                // Analyze the image files
                AnalyzeTests();

                //Display the results
                DisplayResults();
            }

        }

        //Get the data
        private void AnalyzeTests()
        {
            //Get a list of files in the selected folder
            IEnumerable files = Directory.EnumerateFiles(imageDir, "*", SearchOption.AllDirectories);

            //Iterate through list of files
            foreach (string testFile in files)
            {
                FileInfo info = new FileInfo(testFile);

                //If a file is a read-only file, skip the file
                if (!info.Extension.Equals(".smt") &&
                    !info.Extension.Equals(".sat"))
                {
                    continue;
                }

                double size = info.Length/(1024*1024); // in mb
                string ext = info.Extension;
                string title = info.FullName;

                if (size < givenSize)
                    {
                        continue;
                    }
                
                if (!data.ContainsKey(ext))
                    {
                    data.Add(ext, new List<FileData>());
                    }

                FileData f = new FileData(title, size);
                
                data[ext].Add(f);
                
            }
        }

        string textDisplay;
        //Display the results
        private void DisplayResults()
        {

            //Get a list of key values (single characters), sorted alphabetically
            List<string> keys = new List<string>(data.Keys);
            keys.Sort();

            //Display the directory name
            textDisplay = "Big files in: " + imageDir + "\n\n";

            //For each key, display the key, and then its members, also sorted
            foreach (string key in keys)
            {
                textDisplay += key + "\n";

                List<FileData> titles = data[key];
                foreach (FileData title in titles)
                {
                    textDisplay += " \u2022 " + title.ToString() + "\n";
                }
                textDisplay += "\n";
            }

            txtOutput.Text = textDisplay;
        }

        //method to send email to outlook
        public void sendEMailThroughOUTLOOK()
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                string htmlString = textDisplay;
                htmlString = htmlString.Replace("\n", "<p>");
                oMsg.HTMLBody = htmlString;
                
                //Add an attachment.
                //String sDisplayName = "MyAttachment";
                //int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add(@"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName);
                //Subject line
                oMsg.Subject = "Clean these tests.";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("elena.zhelezina@autodesk.com");
                oRecip.Resolve();
                // Send.
                oMsg.Send();

                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error is " + ex.ToString());
            }//end of try block
            
        }//end of Email Method

        private void email_Click(object sender, RoutedEventArgs e)
        {
            sendEMailThroughOUTLOOK();
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            givenSize = Double.Parse(inputSize.Text);

        }

        private void printButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            
            
        }

    }
}
