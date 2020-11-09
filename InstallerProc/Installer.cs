using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace InstallerProc
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);

            Process proc = new Process();


            string TargetPlatform = this.Context.Parameters["AssemblyPath"];
            if (TargetPlatform.Contains("x86"))
            {
                proc.StartInfo.FileName = Path.Combine(System.IO.Path.GetDirectoryName(this.Context.Parameters["AssemblyPath"]), "Installx86.bat");
            }
            else
            {
                proc.StartInfo.FileName = Path.Combine(System.IO.Path.GetDirectoryName(this.Context.Parameters["AssemblyPath"]), "Installx64.bat");
            }

            // 管理者として実行
            proc.StartInfo.Verb = "RunAs";

            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            Process proc = new Process();

            string TargetPlatform = this.Context.Parameters["AssemblyPath"];
            if (TargetPlatform.Contains("x86"))
            {
                proc.StartInfo.FileName = Path.Combine(System.IO.Path.GetDirectoryName(this.Context.Parameters["AssemblyPath"]), "Uninstallx86.bat");
            }
            else
            {
                proc.StartInfo.FileName = Path.Combine(System.IO.Path.GetDirectoryName(this.Context.Parameters["AssemblyPath"]), "Uninstallx64.bat");
            }

            // 管理者として実行
            proc.StartInfo.Verb = "RunAs";

            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }
    }
}
