using EridanSharp;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;

namespace MDS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> filesPath;
        private List<string> FIO;
        private List<string> clientEmails;
        private List<string> senderEmails;
        private List<string> status;
        private List<DateTime> sendTime;
        private string pathToken = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\";
        private string pathAuth = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\";
        private bool wasMerge = false;
        private bool wasSend = false;
        private Marge merge;
        private string senderName;
        private string subject;
        private string bodyMessage;
        //private LoginWindow loginWindow;

        private const string defaultSubject = "Квитанция";
        private const string defaultSenderName = "ГБУ Жилищник района Царицыно";
        private const string defaultBodyMessage =
@"С уважением,
Отдел по работе с юр. и физ. лицами
ГБУ ""Жилищник р-на Царицыно""
ул. Кантемировская, д. 53, к. 1, кабинет № 17
Вход в офис по центру дома, ориентир - желтые перила
тел. 8-499-218-18-05
4992181805@mail.ru
Режим работы:
Понедельник-четверг с 08-00 до 17-00
Пятница с 08-00 до 15-45
Обед с 12-30 до 13-15
";

        public class LogItem
        {
            public string FileName { get; set; }
            public string SenderEmail { get; set; }
            public string ClientEmail { get; set; }
            public string Date { get; set; }
            public string Time { get; set; }
            public string Status { get; set; }
        }

        public MainWindow()
        {
            InitializeComponent();
            filesPath = new List<string>();
            status = new List<string>();
            senderEmails = new List<string>();
            clientEmails = new List<string>();
            FIO = new List<string>();
            sendTime = new List<DateTime>();
            //loginWindow = new LoginWindow();

            tbSubject.Text = subject = defaultSubject;
            tbSenderName.Text = senderName = defaultSenderName;
            rbMessageBody.Text = bodyMessage = defaultBodyMessage;

            // Add columns
            var gridView = new GridView();
            this.lvLog.View = gridView;
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Имя файла",
                Width = 180,
                DisplayMemberBinding = new Binding("FileName")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Почта отправителя",
                Width = 120,
                DisplayMemberBinding = new Binding("SenderEmail")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Почта клиента",
                Width = 120,
                DisplayMemberBinding = new Binding("ClientEmail")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Дата отправки",
                Width = 91,
                DisplayMemberBinding = new Binding("Date")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Время отправки",
                Width = 97,
                DisplayMemberBinding = new Binding("Time")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Статус",
                Width = 91,
                DisplayMemberBinding = new Binding("Status"),

            });
        }

        private void loadingGrid(bool state)
        {
            if (state)
            {
                this.IsEnabled = !state;
                gLoading.Visibility = Visibility.Visible;
                prLoading.IsActive = state;
            }
            else
            {
                this.IsEnabled = !state;
                gLoading.Visibility = Visibility.Hidden;
                prLoading.IsActive = state;
            }
        }
        private void loadingWithAttachmentGrid(bool state)
        {
            if (state)
            {
                this.IsEnabled = !state;
                gLoadingWithAttachments.Visibility = Visibility.Visible;
                prLoadingWithAttachments.IsActive = state;
            }
            else
            {
                this.IsEnabled = !state;
                gLoadingWithAttachments.Visibility = Visibility.Hidden;
                prLoadingWithAttachments.IsActive = state;
            }
        }
        private void ResetUiButtons()
        {
            btnMerge.IsEnabled = true;
            btnSend.IsEnabled = false;
        }
        private void Button_Click_Browse_Excel(object sender, RoutedEventArgs e)
        {
            wasSend = false;
            OpenFileDialog ofdExcel = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), //@"C:\Users\User\Desktop"
                Title = "Browse Excel File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (ofdExcel.ShowDialog() == true)
            {
                tbExcelPath.Text = ofdExcel.FileName;
                ResetUiButtons();
            }
        }

        private void Button_Click_Browse_Folder(object sender, RoutedEventArgs e)
        {
            wasSend = false;
            System.Windows.Forms.FolderBrowserDialog openFileDlg = new System.Windows.Forms.FolderBrowserDialog();
            var result = openFileDlg.ShowDialog();
            if (result.ToString() != string.Empty && result.ToString() != "Cancel")
            {
                tbOutputPath.Text = openFileDlg.SelectedPath;
                ResetUiButtons();
            }
        }

        private void Button_Click_Merge(object sender, RoutedEventArgs e)
        {
            wasSend = false;
            string excelPath = tbExcelPath.Text;
            string outputPath = tbOutputPath.Text;
            bool fillState = false;
            int count = 0;

            if (!string.IsNullOrEmpty(excelPath) && !string.IsNullOrEmpty(outputPath))
            {
                loadingGrid(true);
                fillState = true;
                lvLog.Items.Clear();
            }
            else
            {
                MessageBox.Show("Внесите все необходимые данные для выполнения слияния файлов.", "Предупреждение.", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            if (fillState)
            {
                new Thread(new ThreadStart(
                delegate ()
                {
                    merge = new Marge(excelPath);
                    if (merge.InitializeClients())
                    {
                        try
                        {
                            count = merge.getCount();

                            filesPath = merge.getMargedEW(outputPath);

                            FIO = merge.getFIO();
                            clientEmails = merge.getEmails();

                            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                            new Action(() =>
                            {
                                if (count > 0)
                                {
                                    wasMerge = true;
                                    btnSend.IsEnabled = true;
                                    btnMerge.IsEnabled = false;
                                }
                                loadingGrid(false);
                                MessageBox.Show("Слияние прошло успешно.", "Информация.", MessageBoxButton.OK, MessageBoxImage.Information);
                            })).Wait();
                        }
                        catch (Exception ex)
                        {
                            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                           new Action(() =>
                           {
                               loadingGrid(false);
                               MessageBox.Show(ex.Message, "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error);
                           })).Wait();
                        }
                    }
                    else
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                           new Action(() =>
                           {
                               loadingGrid(false);
                               MessageBox.Show("Некорректная таблица.", "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error);
                           })).Wait();
                        return;
                    }
                }
            )).Start();

            }
        }

        private string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            // Special "url-safe" base64 encode.
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }
        private void Button_Click_Send(object sender, RoutedEventArgs e)
        {
            if (wasMerge)
            {
                loadingWithAttachmentGrid(true);
                string outPath = tbOutputPath.Text;
                Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(
                    delegate ()
                    {
                        int count = merge.getCount();
                        for (int i = 0, j = 0; i < count; i++)
                        {
                            MimeMessage message = new MimeMessage();
                            message.FromName = senderName;
                            message.FromEmail = LoginWindow.oAuth2[j].GetProfile().EmailAddress;
                            message.ToEmail = merge.getEmails()[i];
                            message.ToName = merge.getFIO()[i];
                            message.Subject = subject;
                            message.BodyText = bodyMessage;
                            message.AddAttachment((outPath + "\\" + filesPath[i]).ToString());

                            tbCountSendedMessages.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
                            {
                                tbCountSendedMessages.Text = "Отправлено " + (i + 1) + " из " + count + ".";
                            })).Wait();

                            try
                            {
                                if (i % 5 == 0)
                                {
                                    j++;
                                }
                                if (j > ((int)LoginWindow.oAuth2.Count - 1))
                                {
                                    j = 0;
                                }

                                LoginWindow.oAuth2[j].Send(message);
                                sendTime.Add(DateTime.Now);
                                senderEmails.Add(LoginWindow.oAuth2[j].GetProfile().EmailAddress);
                                status.Add("Успешно");
                            }
                            catch (Exception)
                            {
                                status.Add("Ошибка");
                                sendTime.Add(DateTime.Now);
                                senderEmails.Add("---");
                                Debug.WriteLine("Error: The message wan\'t send.");
                            }

                            lvLog.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
                            {
                                lvLog.Items.Add(new LogItem
                                {
                                    FileName = merge.getFIO()[i] + ".docx",
                                    SenderEmail = LoginWindow.oAuth2[j].GetProfile().EmailAddress,
                                    ClientEmail = merge.getEmails()[i],
                                    Date = DateTime.Now.ToString("dd.MM.yyyy"),
                                    Time = DateTime.Now.ToString("HH:mm:ss"),
                                    Status = status[i]
                                });
                            })).Wait();

                            if (i != (count - 1))
                            {
                                Thread.Sleep(30000);
                            }
                        }

                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                        new Action(() =>
                        {
                            wasMerge = false;
                            wasSend = true;
                            btnSend.IsEnabled = false;
                            btnMerge.IsEnabled = true;
                            loadingWithAttachmentGrid(false);
                        })).Wait();
                    }
                ));
                thread.Start();
            }
        }

        private void Button_Click_Save_Message_Configure(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbSubject.Text) || string.IsNullOrEmpty(rbMessageBody.Text) || string.IsNullOrEmpty(tbSenderName.Text))
            {
                MessageBox.Show("Необходимо заполнить все поля.", "Предупреждение.", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                subject = tbSubject.Text;
                senderName = tbSenderName.Text;
                bodyMessage = rbMessageBody.Text;
                MessageBox.Show("Данные сохранены.", "Информирование.", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void Button_Click_Set_Default(object sender, RoutedEventArgs e)
        {
            tbSubject.Text = defaultSubject;
            tbSenderName.Text = defaultSenderName;
            rbMessageBody.Text = defaultBodyMessage;
        }

        private void Button_Click_Clear_Mail_Struct(object sender, RoutedEventArgs e)
        {
            tbSubject.Text = "";
            tbSenderName.Text = "";
            rbMessageBody.Text = "";
        }

        private void Button_Click_Clear_Main_Form(object sender, RoutedEventArgs e)
        {
            btnMerge.IsEnabled = true;
            btnSend.IsEnabled = false;
            tbSubject.Text = "";
            tbSenderName.Text = "";
            rbMessageBody.Text = "";
        }

        private void Button_Click_Change_Account(object sender, RoutedEventArgs e)
        {
            string pathToken = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\";
            MessageBoxResult result = MessageBox.Show("Вы уверены в том, что хотите удалить данные?", "Вопрос.", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                for (int i = 0; i < 10; i++)
                {
                    if (File.Exists(pathToken + "token_" + i.ToString() + ".json"))
                    {
                        File.Delete(pathToken + "token_" + i.ToString() + ".json");
                    }
                }

                if (File.Exists(pathToken + "authorization_info.json"))
                {
                    File.Delete(pathToken + "authorization_info.json");
                }

                System.Diagnostics.Process.Start(Application.ResourceAssembly.Location);
                Application.Current.Shutdown();
            }
        }

        private void ResetUiMainForm()
        {
            tbExcelPath.Text = "";
            tbOutputPath.Text = "";
            tbLogPath.Text = "";
            lvLog.Items.Clear();
            btnMerge.IsEnabled = true;
            btnSend.IsEnabled = false;
        }

        private void Button_Click_Reset(object sender, RoutedEventArgs e)
        {
            wasMerge = false;
            wasSend = false;
            ResetUiMainForm();
        }

        private void Button_Click_Save_Log(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(tbLogPath.Text))
            {
                string outputLog = tbLogPath.Text;

                Thread thread = new Thread(new ThreadStart(
                    delegate ()
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                            new Action(() =>
                            {
                                loadingGrid(true);
                            }));
                            Marge.SaveLog("template_log3.xlsx", outputLog, merge.getFIO(), senderEmails, merge.getEmails(), sendTime, status);
                        wasSend = false;
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                            new Action(() =>
                            {
                                loadingGrid(false);
                                MessageBox.Show("Отчет успешно сохранен.", "Информация.", MessageBoxButton.OK, MessageBoxImage.Information);
                            }));
                    }
                ));

                if ((lvLog.Items.Count > 0) && (status.Count > 0))
                {
                    thread.Start();
                }
                else
                {
                    MessageBox.Show("Создание отчета невозможно, так как Вы не произвели отправку писем.", "Предупреждение.", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

            }
            else
            {
                MessageBox.Show("Укажите папку для сохранения отчета.", "Предупреждение.", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Button_Click_Browse_Log_Folder(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDlg = new System.Windows.Forms.FolderBrowserDialog();
            var result = openFileDlg.ShowDialog();
            if (result.ToString() != string.Empty && result.ToString() != "Cancel")
            {
                tbLogPath.Text = openFileDlg.SelectedPath;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
