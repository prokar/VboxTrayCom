using Microsoft.Win32;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Management;
using System.ServiceProcess;

namespace VboxTraySvc
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
            this.Committed += new InstallEventHandler(serviceInstaller_Committed);
            this.BeforeUninstall += new InstallEventHandler(serviceInstaller_BeforeUninstall);
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            string userName = Context.Parameters["f1"].Contains("\\") ? Context.Parameters["f1"] : ".\\" + Context.Parameters["f1"];

            serviceProcessInstaller.Username = userName;
            serviceProcessInstaller.Password = Context.Parameters["f2"];

            base.Install(stateSaver);
        }

        private void serviceInstaller_BeforeUninstall(object sender, InstallEventArgs e)
        {
            try
            {
                using (ServiceController sc = new ServiceController(serviceInstaller.ServiceName))
                {
                    if (sc.Status == ServiceControllerStatus.Running)
                    {
                        sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped);
                    }
                }
            }
            catch (InvalidEnumArgumentException)
            {

            }

            foreach (Process p in Process.GetProcessesByName("VboxTray"))
                p.Kill();
        }

        private void serviceInstaller_Committed(object sender, InstallEventArgs e)
        {
            new ServiceController(serviceInstaller.ServiceName).Start();
        }
    }
}