using System.ServiceProcess;
using System.Threading;
//Class References
using VSOMain_NS;




namespace VSO_Monitor_Service_NS
{
    public partial class VSOMonitorService : ServiceBase
    {
        public Thread thread;

        public VSOMonitorService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Thread thread = new Thread(new ThreadStart(VSOMain.startedByService));
            thread.Start();
           
        }

        protected override void OnStop()
        {
            System.Environment.Exit(0);

        }

    }
}

