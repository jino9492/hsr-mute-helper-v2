using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Diagnostics;
using NAudio.CoreAudioApi;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;


namespace HSR_MUTE_HELPER
{

    public partial class Mixer
    {
        #region Find App
        static string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
        static dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

        public static int GetTargetProgramsCount()
        {
            JArray programArray = jsonObject["program"] as JArray;

            if (programArray != null)
                return programArray.Count;
            else
                return 0;
        }

        public static int programCount = GetTargetProgramsCount();

        public static AudioSessionControl[] GetTargetProgram(){
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            MMDevice defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            SessionCollection sessionEnumerator = defaultDevice.AudioSessionManager.Sessions;

            AudioSessionControl[] result = new AudioSessionControl[programCount];
            for (int j = 0; j < programCount; j++)
            {
                for (int i = 0; i < sessionEnumerator.Count; i++)
                {

                    if (Process.GetProcessById((int)sessionEnumerator[i].GetProcessID) != null
                        && Process.GetProcessById((int)sessionEnumerator[i].GetProcessID).ProcessName.Equals((string)jsonObject["program"][j]))
                    {
                        result[j] = sessionEnumerator[i];
                        break;
                    }
                    else
                    {
                        result[j] = null;
                    }
                }
            }

            return result;
        }
        #endregion

        #region Mute App Functions
        public static AudioSessionControl[] target { get; set; } = GetTargetProgram();
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
        static dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

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
                for (int i = 0; i < Mixer.programCount; i++)
                {
                    if (Mixer.target[i] == null || Mixer.target[i].State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateExpired)
                    {
                        UnMuteApp(i);
                        if (Process.GetProcessesByName((string)jsonObject["program"][i]).Length > 0)
                            Mixer.target = Mixer.GetTargetProgram();
                    }
                    else
                    {
                        [DllImport("user32.dll")]
                        static extern IntPtr GetForegroundWindow();

                        IntPtr handle = GetForegroundWindow();
                        try
                        {
                            if (Process.GetProcessById((int)Mixer.target[i].GetProcessID).MainWindowHandle == handle)
                            {
                                UnMuteApp(i);
                            }
                            else
                            {
                                MuteApp(i);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                
            });
            timer.Start();
        }
        #endregion

        #region Mute Function
        public void MuteApp(int idx)
        {
            if (Mixer.target[idx] != null)
            {
                Mixer.target[idx].SimpleAudioVolume.Mute = true;
            }
        }
        public void UnMuteApp(int idx)
        {
            if (Mixer.target[idx] != null)
            {
                Mixer.target[idx].SimpleAudioVolume.Mute = false;
            }
        }
        #endregion
    }
}
