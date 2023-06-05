using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace HSR_MUTE_HELPER
{
    static class Program
    {
        /// <summary>
        /// 해당 애플리케이션의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string processName = Process.GetCurrentProcess().ProcessName;
            Process[] processes = Process.GetProcessesByName(processName);

            bool err = false;

            if (processes.Length > 1)
            {
                MessageBox.Show("이미 프로그램이 실행 중입니다.", "HSR_MUTE_HELPER WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                try
                {
                    string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
                    dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);
                    Icon icon = new Icon(Application.StartupPath + "/Resources/" + (string)jsonObject["icon"]);
                }
                catch
                {
                    MessageBox.Show("실행에 필요한 파일이 일부 누락되었습니다.", "HSR_MUTE_HELPER WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    err = true;
                }

                if (!err)
                {
                    MessageBox.Show("프로그램이 정상적으로 실행되었습니다.", "HSR_MUTE_HELPER INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Run(new HSR_MUTE_HELPER());
                }
            }
        }
    }
}
