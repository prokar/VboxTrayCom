using System;
using System.Diagnostics;
using System.Reflection;
using System.ServiceProcess;
using VirtualBox;

namespace VboxTraySvc
{
    public partial class VboxTraySvc : ServiceBase
    {
        private bool shutdownInProgress = false;
        private bool logging = false;
        private readonly IVirtualBox vbox;

        private const int SERVICE_ACCEPT_PRESHUTDOWN = 0x100;
        private const int SERVICE_CONTROL_PRESHUTDOWN = 0xf;
        private const int VMS_CONTROL_START = 0x82;
        private const int VMS_CONTROL_STOP = 0x83;
        private const int VMS_CONTROL_LOGGING = 0x84;

        public VboxTraySvc()
        {
            InitializeComponent();
            ServiceName = "VboxTraySvc";
            
            AutoLog = false;
            vbox = new VirtualBox.VirtualBox();

            FieldInfo acceptedCommandsFieldInfo =
                    typeof(ServiceBase).GetField("acceptedCommands", BindingFlags.Instance | BindingFlags.NonPublic);
            if (acceptedCommandsFieldInfo == null)
                EventLog.WriteEntry(ServiceName, "acceptedCommands field not found", EventLogEntryType.Information);
            else
            {
                int value = (int)acceptedCommandsFieldInfo.GetValue(this);
                acceptedCommandsFieldInfo.SetValue(this, value | SERVICE_ACCEPT_PRESHUTDOWN);
            }
        }

        protected override void OnCustomCommand(int command)
        {
            if (logging)
                EventLog.WriteEntry(ServiceName, $"OnCustomCommand(int {command}) Fired", EventLogEntryType.Information);

            switch (command)
            {
                case SERVICE_CONTROL_PRESHUTDOWN:
                    shutdownInProgress = true;
                    StopVms();
                    this.Stop();
                    break;
                case VMS_CONTROL_START:
                case VMS_CONTROL_STOP:
                    StartStopVms(command);
                    break;
                case VMS_CONTROL_LOGGING:
                    logging = true;
                    break;
                default:
                    base.OnCustomCommand(command);
                    break;
            }
        }

        protected override void OnStart(string[] args)
        {
            StartStopVms(VMS_CONTROL_START);
        }

        protected override void OnStop()
        {
            if (!shutdownInProgress)
            {
                StopVms();
            }
        }

        private void StartStopVms(int command)
        {
            if (command == VMS_CONTROL_START)
            {
                foreach (IMachine vm in vbox.Machines)
                {
                    Session ses = new Session();
                    IProgress vmProgress;

                    if (vm.AutostartEnabled == 1 && vm.SessionState == SessionState.SessionState_Unlocked)
                    {
                        if (vm.State == MachineState.MachineState_Aborted)
                            return;

                        vmProgress = vm.LaunchVMProcess(ses, "headless", new string[0]);
                        vmProgress.WaitForCompletion(-1);
                        ses.UnlockMachine();
                    }
                }
            }
            else

            if (command == VMS_CONTROL_STOP)
            {
                foreach (IMachine vm in vbox.Machines)
                {
                    Session ses = new Session();
                    IProgress vmProgress;

                    if (vm.AutostopType != AutostopType.AutostopType_Disabled && vm.SessionState == SessionState.SessionState_Locked)
                    {
                        vm.LockMachine(ses, LockType.LockType_Shared);

                        switch (vm.AutostopType)
                        {
                            case AutostopType.AutostopType_AcpiShutdown:
                                ses.Console.PowerButton();

                                DateTime end = DateTime.Now.AddSeconds(10);
                                while (DateTime.Now < end)
                                {
                                    if (vm.SessionState == SessionState.SessionState_Unlocked)
                                        break;
                                }
                                break;

                            case AutostopType.AutostopType_PowerOff:
                                vmProgress = ses.Console.PowerDown();
                                vmProgress.WaitForCompletion(-1);
                                break;

                            case AutostopType.AutostopType_SaveState:
                                vmProgress = ses.Machine.SaveState();
                                vmProgress.WaitForCompletion(-1);
                                break;

                            default:
                                break;
                        }
                    }
                }
            }
            else
                return;
        }

        private void StopVms()
        {
            //Do not use logging when shutdown service !!!
            foreach (IMachine vm in vbox.Machines)
            {
                Session ses = new Session();
                IProgress vmProgress;

                if (vm.AutostopType != AutostopType.AutostopType_Disabled && vm.SessionState == SessionState.SessionState_Locked)
                {
                    switch (vm.AutostopType)
                    {
                        case AutostopType.AutostopType_AcpiShutdown:
                            vm.LockMachine(ses, LockType.LockType_Shared);
                            ses.Console.PowerButton();

                            DateTime end = DateTime.Now.AddSeconds(10);
                            while (DateTime.Now < end)
                            {
                                if (vm.SessionState == SessionState.SessionState_Unlocked)
                                    break;
                            }
                            break;

                        case AutostopType.AutostopType_PowerOff:
                            vm.LockMachine(ses, LockType.LockType_Shared);
                            vmProgress = ses.Console.PowerDown();
                            vmProgress.WaitForCompletion(-1);
                            break;

                        case AutostopType.AutostopType_SaveState:
                            vm.LockMachine(ses, LockType.LockType_Shared);
                            vmProgress = ses.Machine.SaveState();
                            vmProgress.WaitForCompletion(-1);
                            break;

                        default:
                            break;
                    }
                }
            }
        }
    }
}



