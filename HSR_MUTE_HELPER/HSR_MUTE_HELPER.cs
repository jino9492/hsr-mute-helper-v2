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
using System.Threading.Tasks;
using System.Linq;


namespace HSR_MUTE_HELPER
{

    public partial class Mixer
    {
        #region Find App
        static string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
        static dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);
        static JArray programArray = jsonObject["program"] as JArray;
        static List<string> programList = programArray.ToObject<List<string>>();

        public static int GetTargetProgramsCount()
        {
            if (programArray != null)
                return programArray.Count;
            else
                return 0;
        }

        //public static int programCount = GetTargetProgramsCount();

        public static Dictionary<Process, AudioSessionControl> GetTargetProgram(){
            try
            {
                MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
                MMDevice defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                SessionCollection sessionEnumerator = defaultDevice.AudioSessionManager.Sessions;

                Dictionary<Process, AudioSessionControl> result = new Dictionary<Process, AudioSessionControl>();
                for (int i = 0; i < sessionEnumerator.Count; i++)
                {

                    if (Process.GetProcessById((int)sessionEnumerator[i].GetProcessID) != null)
                    {
                        int pid = (int)sessionEnumerator[i].GetProcessID;
                        Process process = Process.GetProcessById(pid);
                        if (programList.Contains(process.ProcessName))
                        {
                            result.Add(process, sessionEnumerator[i]);
                        }
                        else
                        {
                            try
                            {
                                result.Remove(process);
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }

                return result;
            }
            catch {
                return GetTargetProgram();
            }
        }
        #endregion

        #region Mute App Functions
        public static Dictionary<Process, AudioSessionControl> target { get; set; } = GetTargetProgram();
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
            FindTarget();

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
        private Timer timer2;
        public bool isActivatingCTS;
        
        public void FindTarget()
        {
            ChangeTargetState().Wait();
            timer2 = new Timer();
            timer2.Interval = 5000;
            timer2.Tick += new EventHandler(async (sender, e) =>
            {
                Mixer.target = await Task.Run(() => Mixer.GetTargetProgram());
                if (!isActivatingCTS)
                {
                    await ChangeTargetState();
                }
            });

            timer2.Start();
        }

        public async Task ChangeTargetState()
        {
            isActivatingCTS = true;
            IntPtr lastHandle = new IntPtr();
            IntPtr handle = new IntPtr();
            
            timer = new Timer();
            timer.Interval = 16;
            timer.Tick += new EventHandler(async (sender, e) => {
                    await Task.Run(() => {
                    foreach(KeyValuePair<Process, AudioSessionControl> t in Mixer.target)
                    {
                        try
                        {
                            [DllImport("user32.dll")]
                            static extern IntPtr GetForegroundWindow();

                            handle = GetForegroundWindow();

                            if (handle != lastHandle)
                            {
                                if (t.Key.MainWindowHandle == handle)
                                {
                                    UnMuteApp(t.Value);
                                }
                                else
                                {
                                    MuteApp(t.Value);
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                        
                    }
                    
                    lastHandle = handle;

                    if (Mixer.target.Count <= 0)
                    {
                        isActivatingCTS = false;
                        timer.Stop();
                    }
                            
                });
            });
            timer.Start();
        }
        #endregion

        #region Mute Function
        public void MuteApp(AudioSessionControl target)
        {
            if (target != null)
            {
                target.SimpleAudioVolume.Mute = true;
            }
        }
        public void UnMuteApp(AudioSessionControl target)
        {
            if (target != null)
            {
                target.SimpleAudioVolume.Mute = false;
            }
        }
        #endregion
    }
}
