using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;


public class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        new Form1();
        Application.Run();
    }
}

public class Form1 : Form
{
    private NotifyIcon ico;
    private string appDir;

    public Form1()
    {
        appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        // タスクバーに表示しない
        this.ShowInTaskbar = false;
        setComponents();
        Task.Run(() => checkWU());
    }

    private void setComponents()
    {
        // タスクトレイアイコンの設定
        ico = new NotifyIcon();
        ico.Icon = new Icon(appDir + @"\wmploc_132_9.ico");
        ico.Visible = true;
        ico.Text = "Windowsアップデートチェッカー - ISHII_Tools";

        // コンテキストメニュー
        ContextMenuStrip menu = new ContextMenuStrip();
        ToolStripMenuItem item = new ToolStripMenuItem();
        item.Text = "終了";
        item.Click += ((object sender, EventArgs e) => { ico.Dispose(); Application.Exit(); });
        menu.Items.Add(item);
        ico.ContextMenuStrip = menu;
    }

    private void checkWU()
    {
        // ツールを終了するまで監視し続ける
        while (true)
        {
            const string BASE = @"SOFTWARE\Microsoft\Windows\CurrentVersion";

            bool RebootPending = existsKey(BASE + @"\Component Based Servicing\RebootPending");
            bool RebootRequired = existsKey(BASE + @"\WindowsUpdate\Auto Update\RebootRequired");

            if (RebootPending || RebootRequired)
            {
                string title = "本日Windowsアップデートあり";
                string message = "シャットダウン時に更新が入る可能性があります。\n休憩時間等に再起動をおすすめします。";
                notify(title, message);
            }
            const int time = 3600 * 1000;
            Thread.Sleep(time);
        }
    }

    private bool existsKey(string key)
    {
        try
        {
            Registry.LocalMachine.OpenSubKey(key);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void notify(string title, string msg)
    {
        ico.BalloonTipTitle = title;
        ico.BalloonTipText = msg;
        ico.ShowBalloonTip(3000);
        MessageBox.Show(title + "\n\n" + msg, "Windowsアップデートチェッカー - ISHII_Tools");
    }
}