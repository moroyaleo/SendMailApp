using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SendMailApp
{
    /// <summary>
    /// ConfigWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ConfigWindow : Window
    {
        public ConfigWindow()
        {
            InitializeComponent();
        }

        private void btDefault_Click(object sender, RoutedEventArgs e)
        {
            Config cf = Config.GetInstance().getDefaultStatus();

            tbSmtp.Text = cf.Smtp;
            tbPort.Text = cf.Port.ToString();
            tbUserName.Text = cf.MailAddress;
            tbPassWord.Password = cf.PassWord;
            cbSsl.IsChecked = cf.Ssl;
            tbSender.Text = cf.MailAddress;


        }

        //更新
        private void btApply_Click(object sender, RoutedEventArgs e)
        {
            Warning();
            Config.GetInstance().UpdateStatus(
                tbSmtp.Text,
                tbUserName.Text,
                tbPassWord.Password,
                int.Parse(tbPort.Text),
                cbSsl.IsChecked ?? false);
            
        }

        //OKボタン
        private void btOk_Click(object sender, RoutedEventArgs e)
        {
            Warning();
            btApply_Click(sender, e);
            

        }

        public void Warning()
        {
            if (tbSmtp.Text == "")
            {
                MessageBox.Show("Smtpを入力してください");
            }
            else if (tbUserName.Text == "")
            {
                MessageBox.Show("ユーザー名を入力してください");
            }
            else if (int.Parse(tbPort.Text) == 0)
            {
                MessageBox.Show("ポート番号を入力してください");
            }
            else if (tbPassWord.Password == "")
            {
                MessageBox.Show("パスワードを入力してください");
            }
            else if (tbSender.Text == "")
            {
                MessageBox.Show("送信元を入力してください");
            }
            this.Close();
        }

        //キャンセルボタン
        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            Config ctf = Config.GetInstance();
            if (tbSmtp.Text != ctf.Smtp || int.Parse(tbPort.Text) != ctf.Port || 
                tbUserName.Text != ctf.MailAddress || tbPassWord.Password != ctf.PassWord
                || cbSsl.IsChecked != ctf.Ssl || tbSender.Text != ctf.MailAddress)
            {
                MessageBoxResult result = MessageBox.Show("変更が反映されていません", "警告",
                    MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                if (result == MessageBoxResult.OK)
                {
                    this.Close();
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    
                }

            }

            
        }

        //ロード時に一度だけ呼び出される
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Config ctf = Config.GetInstance();

            tbSmtp.Text = ctf.Smtp;
            tbPort.Text = ctf.Port.ToString();
            tbUserName.Text = ctf.MailAddress;
            tbPassWord.Password = ctf.PassWord;
            cbSsl.IsChecked = ctf.Ssl;
            tbSender.Text = ctf.MailAddress;
        }
    }
}
