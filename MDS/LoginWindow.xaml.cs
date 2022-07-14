using EridanSharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Threading;

namespace MDS
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        //private AuthorizationInfo authorizationInfo;
        public static List<OAuth2> oAuth2;
        private CountAccountsPage countAccountsPage;
        private SignInPage signInPage;
        private AuthLoadingPage authLoadingPage;
        private bool stateAuth = false;
        private Thread tAuth;
        private string pathToken;
        //static public OAuth2 oAuth2;
        public LoginWindow()
        {
            InitializeComponent();

            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\");
            }

            //authorizationInfo = new AuthorizationInfo();
            oAuth2 = new List<OAuth2>();
            MainAuthFrame.Content = new CountAccountsPage();
            countAccountsPage = new CountAccountsPage();
            signInPage = new SignInPage();
            authLoadingPage = new AuthLoadingPage();

            MainAuthFrame.Content = countAccountsPage;
            countAccountsPage.btnExit.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnExit_Click));
            authLoadingPage.btnExit.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnExit_Click));
            countAccountsPage.btnCon.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnContinue_Click));
            countAccountsPage.btnClear.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnClear_Click));
            signInPage.btnGoogleSignIn.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnGoogleSignIn_Click));
            signInPage.btnBack.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(btnBack_Click));

            for (int i = 0; i < 10; i++)
            {
                countAccountsPage.cbAccountCount.Items.Add((i + 1).ToString());
            }

            countAccountsPage.cbService.Items.Add("Google");
            //cbService.Items.Add("Yahoo");
            //cbService.Items.Add("MailRu");
            //cbService.Items.Add("Rambler");

            pathToken = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\";
            string authInfoPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\" + "authorization_info.json";
            if (File.Exists(authInfoPath))
            {
                Dictionary<string, string> authInfo = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(authInfoPath));

                for (int i = 0; i < Int32.Parse(authInfo["Count"]); i++)
                {
                    // Google OAuth2 Initialize
                    authLoadingPage.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                    new Action(() =>
                    {
                        authLoadingPage.tblAccountIndex.Text = "Вход в аккаунт №" + (i + 1).ToString() + ".";
                    })).Wait();
                    oAuth2.Add(new OAuth2(
                        "271195115551-ceelr2teuhbeq3guse4edk13pcjapaea.apps.googleusercontent.com",
                        "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",
                        "Resources\\AuthPages\\success_auth.html",
                        "Resources\\AuthPages\\unsuccess_auth.html",
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\token_" + i.ToString() + ".json"
                    ));

                    if (IsConnectedToInternet())
                    {
                        switch (authInfo["CurrentService"])
                        {
                            case AuthorizationInfo.Services.Google:

                                if (oAuth2[i].CheckExistToken())
                                {
                                    if (oAuth2[i].InitializeProfile())
                                    {
                                        stateAuth = true;
                                    }
                                    else
                                    {
                                        stateAuth = false;
                                        break;
                                    }
                                }
                                else
                                {
                                    stateAuth = false;
                                    break;
                                }
                                break;

                            case AuthorizationInfo.Services.Yahoo:

                                break;

                            case AuthorizationInfo.Services.MailRu:

                                break;

                            case AuthorizationInfo.Services.Rambler:

                                break;
                        }
                    }
                    else
                    {
                        stateAuth = false;
                        MessageBox.Show("Нет интернет соединения.", "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                        this.Close();
                    }
                }
            }



            if (stateAuth)
            {
                Application.Current.MainWindow = new MainWindow();
                Application.Current.MainWindow.Topmost = true;
                Application.Current.MainWindow.Show();
                Application.Current.MainWindow.Topmost = false;
                this.Close();
            }
        }

        private void StartAuthentication()
        {
            int count = 0;
            string currentService = "";
            bool stateAuthThroughBrowser = false;

            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                new Action(() =>
                {
                    MainAuthFrame.Content = authLoadingPage;
                    count = Int32.Parse(countAccountsPage.cbAccountCount.SelectedItem.ToString());
                    currentService = countAccountsPage.cbService.SelectedItem.ToString();

                })).Wait();
            //tAuth = new Thread(()=> {
            for (int i = 0; i < count; i++)
            {
                // Google OAuth2 Initialize
                authLoadingPage.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                new Action(() =>
                {
                    authLoadingPage.tblAccountIndex.Text = "Вход в аккаунт №" + (i + 1).ToString() + ".";
                })).Wait();

                oAuth2.Add(new OAuth2(
                    "271195115551-ceelr2teuhbeq3guse4edk13pcjapaea.apps.googleusercontent.com",
                    "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",
                    "Resources\\AuthPages\\success_auth.html",
                    "Resources\\AuthPages\\unsuccess_auth.html",
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\token_" + i.ToString() + ".json"
                ));

                if (IsConnectedToInternet())
                {
                    switch (currentService)
                    {
                        case AuthorizationInfo.Services.Google:

                            if (oAuth2[i].CheckExistToken())
                            {
                                if (oAuth2[i].InitializeProfile())
                                {
                                    stateAuth = true;
                                }
                                else
                                {
                                    stateAuth = false;
                                    break;
                                }
                            }
                            else
                            {
                                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                                    new Action(() =>
                                    {
                                        authLoadingPage.tblInfo.Text = "Ожидание.";
                                    })).Wait();

                                try
                                {
                                    if (!oAuth2[i].Authentication())
                                    {
                                        stateAuthThroughBrowser = false;
                                        stateAuth = false;
                                        break;
                                    }
                                    else
                                    {
                                        oAuth2[i].InitializeProfile();
                                        stateAuthThroughBrowser = true;
                                        stateAuth = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                                }
                                
                                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                                    new Action(() =>
                                    {
                                        authLoadingPage.tblInfo.Text = "Пожалуйста, подождите.";
                                    })).Wait();
                            }
                            break;

                        case AuthorizationInfo.Services.Yahoo:

                            break;

                        case AuthorizationInfo.Services.MailRu:

                            break;

                        case AuthorizationInfo.Services.Rambler:

                            break;
                    }
                }
                else
                {
                    stateAuth = false;
                    MessageBox.Show("Нет интернет соединения.", "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    this.Close();
                }
            }
            if (stateAuthThroughBrowser)
            {
                oAuth2[0].ShowSucessPage();
            }
            else if (!stateAuthThroughBrowser && !stateAuth)
            {
                oAuth2[0].ShowUnsucessPage();
            }
        //});

            //tAuth.Start();
            //tAuth.Join();

            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal,
                new Action(() =>
                {
                    if (stateAuth)
                    {
                        Application.Current.MainWindow = new MainWindow();
                        Application.Current.MainWindow.Topmost = true;
                        Application.Current.MainWindow.Show();
                        Application.Current.MainWindow.Topmost = false;
                        JObject authInfo = new JObject(
                            new JProperty("Count", count),
                            new JProperty("CurrentService", currentService));

                        // Create the file, or overwrite if the file exists.
                        FileStream fs = File.Create(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\" + "authorization_info.json");
                        fs.Close();
                        File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\LostSkyCorp\\data\\" + "authorization_info.json", authInfo.ToString());
                        //this.Visibility = Visibility.Hidden;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Не удалось пройти аутентификацию.", "Ошибка.", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                        //this.Close();
                    }
                })).Wait();
        }
        public static bool IsConnectedToInternet()
        {
            try
            {
                using (var client = new WebClient())
                using (client.OpenRead("https://yandex.ru/"))
                    return true;
            }
            catch
            {
                return false;
            }
        }

        private void btnContinue_Click(object sender, RoutedEventArgs e)
        {
            if (countAccountsPage.cbAccountCount.SelectedIndex >= 0 && countAccountsPage.cbService.SelectedIndex >= 0)
            {
                MainAuthFrame.Content = signInPage;
            }
            else
            {
                MessageBox.Show("Необходимо заполнить все поля.", "Предупреждение.", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
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
            }
        }
        private void btnGoogleSignIn_Click(object sender, RoutedEventArgs e)
        {
            tAuth = new Thread(() =>
            {
                StartAuthentication();
            });

            tAuth.Start();
            //tAuth.Join();

        }
        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            MainAuthFrame.Content = countAccountsPage;
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

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            if (tAuth != null)
            {
                if (tAuth.IsAlive)
                {
                    tAuth.Abort();
                    if (File.Exists(pathToken + "authorization_info.json"))
                    {
                        File.Delete(pathToken + "authorization_info.json");
                    }
                    for (int i = 0; i < 10; i++)
                    {
                        if (File.Exists(pathToken + "token_" + i.ToString() + ".json"))
                        {
                            File.Delete(pathToken + "token_" + i.ToString() + ".json");
                        }
                    }
                }
            }
            Environment.Exit(0);
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }
    }
}
