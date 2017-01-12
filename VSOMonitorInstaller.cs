using System.ComponentModel;
using System.ServiceProcess;

namespace VSO_Monitor
{
    [RunInstaller(true)]
    public partial class VSOMonitorInstaller : System.Configuration.Install.Installer
    {
        public VSOMonitorInstaller()
        {
            // InitializeComponent();
            serviceProcessInstaller1 = new ServiceProcessInstaller();
            serviceProcessInstaller1.Account = ServiceAccount.LocalSystem;
            serviceInstaller1 = new ServiceInstaller();
            serviceInstaller1.ServiceName = "VSO Monitor Service";
            serviceInstaller1.DisplayName = "VSO Monitor Service";
            serviceInstaller1.Description = "Used for monitoring changes in VSO.";
            serviceInstaller1.StartType = ServiceStartMode.Automatic;
            Installers.Add(serviceProcessInstaller1);
            Installers.Add(serviceInstaller1);

        }

    }
}
