using System.ServiceProcess;

namespace VboxTraySvc
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
#if DEBUG
#else
#endif

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new VboxTraySvc()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
