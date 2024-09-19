using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Timer = System.Timers.Timer;
using System.Timers;


namespace GoogleDriveFileUploaded
{
    public partial class Form1 : Form
    {
        // string credentialPath = "C:\\Users\\Admin\\Desktop\\Maverick Services\\GoogleDriveFileUploaded\\GoogleDriveFileUploaded\\bin\\Debug\\Credentials.json";
        string credentialPath = "C:\\Users\\Admin\\Desktop\\Maverick Services\\GoogleDriveFileUploaded\\GoogleDriveFileUploaded\\bin\\Debug\\creditial.json";

        string folderId = "1x5ECioJYbygeHUmY5ZwKHaCXDqEt5Qph";
        //iqtd jimn piag zymt
        //radhakdam73@gmail.com
        //bzxj qgzp ghnk ejay
        //serviceaccoradha@boreal-totality-434813-i7.iam.gserviceaccount.com
        string FileToUpload = "C:\\Users\\Admin\\Desktop\\Maverick Services\\GoogleDriveFileUploaded\\GoogleDriveFileUploaded\\files\\Maverick.accdb";
        private static Stream fileStream;
        private readonly Timer emailTimer;


        public Form1()
        {
            InitializeComponent();
            AddTextBoxFocusEvents(this);
            textBoxAttached.Click += textBoxAttached_TextChanged;
            emailTimer = new Timer(120000);
            emailTimer.Elapsed += OnTimedEvent;  // Attach event handler
            emailTimer.AutoReset = true;         // Enable repetition
            emailTimer.Enabled = true;

            if (File.Exists(FileToUpload))
            {
                UploadFileToGoogleDriveAutomatically(credentialPath, folderId, FileToUpload);
            }
            else
            {
                MessageBox.Show($"The file '{FileToUpload}' does not exist.");
            }
        
    }
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            SendEmail2minteveyWithAttachment();
        }

        private void SendEmail2minteveyWithAttachment()
        {
            try
            {
                var fileLink = @"https://drive.google.com/drive/folders/1x5ECioJYbygeHUmY5ZwKHaCXDqEt5Qph?usp=sharing";
                string fromEmail = "radhakadam73@gmail.com";
                string fromPassword = "iqtd jimn piag zymt"; // Use the App Password here
                string toEmail = "tejasshingavi9999@gmail.com";
                string subject = "QR Code Payment";
                string body = $"Hi Sir/Mam,\r\nWe have generated your payment QR code. You can download it using the following link: {fileLink}\r\nPlease scan it with your UPI app to complete the transaction.\r\nThank you for your prompt attention to this matter.\r\nWarm regards,\r\nRadha Kadam";

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(fromEmail);
                mail.To.Add(toEmail);
                mail.Subject = subject;
                mail.Body = body;

                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
                smtpClient.EnableSsl = true;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(fromEmail, fromPassword);
                smtpClient.Send(mail);

                MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (SmtpException smtpEx)
            {
                MessageBox.Show("SMTP Error: " + smtpEx.Message + "\nCheck your SMTP settings and credentials.", "SMTP Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void UploadFileToGoogleDriveAutomatically(string credentialPath, string folderId, string FileToUpload)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(new[] { DriveService.ScopeConstants.DriveFile });

                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Drive"
                });

                var metadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = Path.GetFileName(FileToUpload),
                    Parents = new List<string> { folderId }
                };

                FilesResource.CreateMediaUpload request;
                using (var fileStream = new FileStream(FileToUpload, FileMode.Open))
                {
                    request = service.Files.Create(metadata, fileStream, GetMimeTypes(FileToUpload));
                    request.Fields = "id";
                    request.Upload();
                }

                var uploadedFile = request.ResponseBody;
                Console.WriteLine($"File '{metadata.Name}' uploaded with ID: {uploadedFile.Id}");
            }
        }

        private static string GetMimeTypes(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }
        // Adds focus event to all TextBox controls in the form
        private void AddTextBoxFocusEvents(Control control)
        {
            foreach (Control c in control.Controls)
            {
                if (c is TextBox textBox)
                {
                    textBox.Enter -= TextBox_Enter;
                    textBox.Leave -= TextBox_Leave;

                    textBox.Enter += TextBox_Enter;
                    textBox.Leave += TextBox_Leave;
                }

                if (c.Controls.Count > 0)
                {
                    AddTextBoxFocusEvents(c);
                }
            }
        }

        // Change background color on TextBox enter
        private void TextBox_Enter(object sender, EventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.BackColor = Color.Black;
                textBox.ForeColor = Color.White;
            }
        }

        // Revert background color on TextBox leave
        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.BackColor = Color.White;
                textBox.ForeColor = Color.Black;
            }
        }

        // Enable pressing Enter to switch between controls
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                if (this.ActiveControl is TextBox || this.ActiveControl is ComboBox || this.ActiveControl is DateTimePicker)
                {
                    this.SelectNextControl(this.ActiveControl, true, true, true, true);
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // Button click event to upload file and send email
        private void button1_Click(object sender, EventArgs e)
        {
            string emailServer = textBoxserver.Text;
            string serverAddress = textBoxAddress.Text;
            string fromEmail = textBoxFromID.Text;
            string password = textBoxPass.Text;
            string toEmail = textBoxToID.Text;
            string filePath = textBoxAttached.Text;
            string subject = textsub.Text;
            string additionalText = textBoxtxt.Text;

            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(toEmail))
            {
                MessageBox.Show("Please provide both file path and recipient email address.");
                return;
            }

            try
            {
                // Upload the file to Google Drive
                UploadFileToGoogleDrive(credentialPath, folderId, filePath);

                // After uploading, send an email
              // SendEmail(emailServer, serverAddress, fromEmail, password, toEmail, subject, additionalText, filePath);
                MessageBox.Show("File uploaded to Google Drive and email sent successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        //private void SendEmail(string emailServer, string serverAddress, string fromEmail, string password, string toEmail, string subject, string additionalText, string filePath)
        //{
        //    try
        //    {
        //        MailMessage mail = new MailMessage(fromEmail, toEmail)
        //        {
        //            Subject = subject,
        //            Body = additionalText
        //        };

        //        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        //        {
        //            Attachment attachment = new Attachment(filePath);
        //            mail.Attachments.Add(attachment);
        //        }

        //        SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587)
        //        {
        //            Credentials = new NetworkCredential(fromEmail, password), // Use App Password here
        //            EnableSsl = true,
        //            DeliveryMethod = SmtpDeliveryMethod.Network,
        //            UseDefaultCredentials = false
        //        };

        //        smtpClient.Send(mail); // For async: smtpClient.SendAsync(mail, null);

        //        MessageBox.Show("Email sent successfully!");
        //    }
        //    catch (SmtpException smtpEx)
        //    {
        //        MessageBox.Show("SMTP Error: " + smtpEx.Message);
        //        Console.WriteLine("SMTP Error: " + smtpEx.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error sending email: " + ex.Message);
        //        Console.WriteLine("General Error: " + ex.ToString());
        //    }
        //}



        static void UploadFileToGoogleDrive(string credentialPath, string folderId, string filePath)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(new[]
                {
            DriveService.ScopeConstants.DriveFile
        });

                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Drive"
                });

                var metadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = Path.GetFileName(filePath),
                    Parents = new List<string> { folderId }
                };

                FilesResource.CreateMediaUpload request;
                using (var fileStream = new FileStream(filePath, FileMode.Open))
                {
                    request = service.Files.Create(metadata, fileStream, GetMimeType(filePath));
                    request.Fields = "id";
                    request.Upload();
                }

                var uploadedFile = request.ResponseBody;
                Console.WriteLine($"File '{metadata.Name}' uploaded with ID: {uploadedFile.Id}");
            }
        }

        static string GetMimeType(string filePath)
        {
            // A method to determine MIME type based on file extension
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            switch (extension)
            {
                case ".jpg": case ".jpeg": return "image/jpeg";
                case ".png": return "image/png";
                case ".pdf": return "application/pdf";
                // Add more types as needed
                default: return "application/octet-stream";
            }
        }

        // Exit button click event
        private void btnexite_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // TextBox change event to open file dialog for file attachment
        // Flag to control whether the file dialog should open
        private bool isDialogOpen = false;

        private void textBoxAttached_TextChanged(object sender, EventArgs e)
        {
            if (!isDialogOpen) // Check if the dialog is already open
            {
                isDialogOpen = true; // Set flag to true to prevent reopening

                using (OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
                    Title = "Select a PDF file"
                })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        textBoxAttached.Text = openFileDialog.FileName;
                    }
                }

                isDialogOpen = false; // Reset the flag to allow future openings
            }
        }

    }
}
