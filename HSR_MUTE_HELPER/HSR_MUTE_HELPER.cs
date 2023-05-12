using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using NAudio.CoreAudioApi;
using Newtonsoft.Json;
using System.Runtime.InteropServices;


namespace HSR_MUTE_HELPER
{

    public partial class Mixer
    {
        #region Find App
        public static AudioSessionControl GetTargetProgram(){
            string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
            dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            MMDevice defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            SessionCollection sessionEnumerator = defaultDevice.AudioSessionManager.Sessions;

            for (int i = 0; i < sessionEnumerator.Count; i++)
            {
                if (Process.GetProcessById((int)sessionEnumerator[i].GetProcessID) != null 
                    && Process.GetProcessById((int)sessionEnumerator[i].GetProcessID).ProcessName.Equals((string)jsonObject["program"], StringComparison.OrdinalIgnoreCase))
                {
                    return sessionEnumerator[i];
                }
            }

            return null;
        }
        #endregion

        #region Mute App Functions
        public static AudioSessionControl target { get; set; } = GetTargetProgram();
        #endregion
    }

    public partial class HSR_MUTE_HELPER : Form
    {
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80;
                return cp;
            }
        }

        static string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
        dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

        private ContextMenu contextMenu;
        private NotifyIcon notifyIcon;

        public HSR_MUTE_HELPER()
        {
            CheckTargetState();

            contextMenu = new ContextMenu();
            contextMenu.MenuItems.Add(new MenuItem("Exit", new EventHandler((sender, e) =>
            {
                notifyIcon.Visible = false;
                Application.Exit();
            })));

            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = new Icon(Application.StartupPath + "/Resources/" + (string)jsonObject["icon"]);
            notifyIcon.ContextMenu = contextMenu;
            notifyIcon.Visible = true;

            this.BackColor = Color.Magenta;
            this.TransparencyKey = Color.Magenta;
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.ShowInTaskbar = false;

        }

        #region Target Program State Check Function
        private Timer timer;
        public void CheckTargetState()
        {
            
            timer = new Timer();
            timer.Interval = 16;
            timer.Tick += new EventHandler((sender, e) => {
                if(Mixer.target == null
                || Mixer.target.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateExpired)
                {
                    Application.Exit();
                }

                [DllImport("user32.dll")]
                static extern IntPtr GetForegroundWindow();

                IntPtr handle = GetForegroundWindow();
                Debug.WriteLine(handle);
                if (Process.GetProcessById((int)Mixer.target.GetProcessID).MainWindowHandle == handle)
                {
                    UnMuteApp((string)jsonObject["program"]);
                }
                else
                {
                    MuteApp((string)jsonObject["program"]);
                }
            });
            timer.Start();
        }
        #endregion

        #region Mute Function
        public void MuteApp(string programName)
        {
            if (Mixer.target != null)
            {
                Mixer.target.SimpleAudioVolume.Mute = true;
            }
        }
        public void UnMuteApp(string programName)
        {
            if (Mixer.target != null)
            {
                Mixer.target.SimpleAudioVolume.Mute = false;
            }
        }
        #endregion
    }
}
