using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using VirtualBox;

namespace VboxTray
{
    public class MyApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private ToolStripMenuItem menuItem, subMenuItem;
        private readonly IContainer components;
        private readonly IVirtualBox vbox;

        public MyApplicationContext()
        {            
            vbox = new VirtualBox.VirtualBox();

            components = new Container();
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            trayIcon = new NotifyIcon(components)
            {
                Icon = Properties.Resources.VirtualBox_vbox,
                Visible = true,
                Text = "VboxTray"
            };

            trayIcon.ContextMenuStrip = new ContextMenuStrip();
            trayIcon.ContextMenuStrip.Opening += new CancelEventHandler(OnOpening);
            RebuildMenu();
        }

        private void OnOpening(object sender, CancelEventArgs e)
        {
            RebuildMenu();
            e.Cancel = false;
        }

        private void RebuildMenu()
        {
            trayIcon.ContextMenuStrip.SuspendLayout();
            trayIcon.ContextMenuStrip.Items.Clear();

            ToolStripLabel mainLabel = new ToolStripLabel
            {
                Name = "MainLabel",
                Text = "Vms List",
                Margin = new Padding(50, 3, 3, 3),
                Enabled = false,
                AutoToolTip = false,
                ToolTipText = "Double Click to trigger Running/SaveState selected Vm NOW."
            };

            mainLabel.Font = new Font(mainLabel.Font, FontStyle.Bold);
            trayIcon.ContextMenuStrip.Items.Add(mainLabel);
            trayIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            foreach (IMachine vm in vbox.Machines)
            {
                var str = ((VmState)vm.State).ToString();

                menuItem = new ToolStripMenuItem
                {
                    Name = $"{vm.Name}_MenuItem",
                    Tag = vm.Name,
                    Text = $"{vm.Name} - {str}",
                    Image = SetImg(str)
                };

                menuItem.DoubleClickEnabled = true;
                menuItem.DoubleClick += new EventHandler(MenuItemDoubleClick);
                trayIcon.ContextMenuStrip.Items.Add(menuItem);

                bool active = false;
                bool trigger = false;
                ToolStripLabel subLabel = new ToolStripLabel
                {
                    Name = "SubLabel_" + menuItem.Tag,
                    Text = $"Auto shutdown type{Environment.NewLine}for AutoStartEnabled",
                    Margin = new Padding(3, 3, 3, 3),
                    Enabled = false,
                    AutoToolTip = false,
                    ToolTipText = $"Click to change Shutdown Type for Auto Starting Vm." +
                                    $"{Environment.NewLine}" +
                                    $"If AutoStartDisabled - Autostart None."
                };
                subLabel.Font = new Font(mainLabel.Font, FontStyle.Bold);

                menuItem.DropDownItems.Add(subLabel);
                menuItem.DropDownItems.Add(new ToolStripSeparator());

                subMenuItem = new ToolStripMenuItem();

                foreach (AutostopType state in Enum.GetValues(typeof(AutostopType)))
                {
                    active = false;
                    if (state != AutostopType.AutostopType_Disabled)
                    {
                        if (vm.AutostopType == state)
                        {
                            active = true;
                            trigger = true;
                        }
                        subMenuItem = AddSubMenu(menuItem, ((VmShutdownType)state).ToString(), active);
                        menuItem.DropDownItems.Add(subMenuItem);
                        active = false;
                    }
                }

                menuItem.DropDownItems.Add(new ToolStripSeparator());
                subMenuItem = AddSubMenu(menuItem, (VmShutdownType.AutoStartDisabled).ToString(), !trigger);
                menuItem.DropDownItems.Add(subMenuItem);
                trigger = false;
            }


            menuItem = new ToolStripMenuItem
            {
                Name = "CloseMenuItem",
                Text = "Close the VboxTray program"
            };

            menuItem.Click += new EventHandler(CloseMenuItem_Click);
            trayIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            trayIcon.ContextMenuStrip.Items.Add(menuItem);

            trayIcon.ContextMenuStrip.ResumeLayout(true);
        }

        private ToolStripMenuItem AddSubMenu(ToolStripMenuItem item, string state, bool active = false)
        {
            ToolStripMenuItem menu = new ToolStripMenuItem()
            {
                Name = $"{item.Name}_{state}",
                Text = state,
                Tag = item.Tag
            };

            if (active)
                menu.Font = new Font(menu.Font, FontStyle.Bold);

            menu.Click += new EventHandler(VmAutoStopTypeChange);

            return menu;
        }

        private void VmAutoStopTypeChange(object sender, EventArgs e)
        {
            var obj = (ToolStripMenuItem)sender;
            IMachine vm = vbox.FindMachine(obj.Tag.ToString());
            Session ses = new Session();

            if (obj.Text.ToString().Equals((VmShutdownType.AutoStartDisabled).ToString()))
            {
                vm.LockMachine(ses, LockType.LockType_Shared);
                ses.Machine.AutostartEnabled = 0;
                ses.Machine.AutostopType = AutostopType.AutostopType_Disabled;
                ses.Machine.SaveSettings();
                ses.UnlockMachine();
            }
            else
            {
                vm.LockMachine(ses, LockType.LockType_Shared);
                ses.Machine.AutostartEnabled = 1;
                var enu = (VmShutdownType)Enum.Parse(typeof(VmShutdownType), obj.Text);
                ses.Machine.AutostopType = (AutostopType)enu;
                ses.Machine.SaveSettings();
                ses.UnlockMachine();
            }
            ses = null;
        }

        private void MenuItemDoubleClick(object sender, EventArgs e)
        {
            var obj = (ToolStripMenuItem)sender;

            Session ses = new Session();
            IMachine vm = vbox.FindMachine(obj.Tag.ToString());

            if (vm.SessionState == SessionState.SessionState_Locked)
            {
                vm.LockMachine(ses, LockType.LockType_Shared);
                IProgress vmProgress = ses.Machine.SaveState();
                CatchStatus(obj, vmProgress, vm);
                if (vm.SessionState == SessionState.SessionState_Locked)
                    ses.UnlockMachine();
            }
            else
            {
                IProgress vmProgress = vm.LaunchVMProcess(ses, "headless", new string[0]);
                CatchStatus(obj, vmProgress, vm);
                if (vm.SessionState == SessionState.SessionState_Locked)
                    ses.UnlockMachine();
            }
        }

        private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you really want to close me?",
                    "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                Dispose(true);
                Application.Exit();
            }
        }

        public Image SetImg(string name)
        {
            Image image = (Image)Properties.Resources.ResourceManager.GetObject(
                            name: $"state_{name.ToLower()}_16px_x2",
                            culture: CultureInfo.CurrentCulture
                            );
            return image;
        }

        public enum VmShutdownType
        {
            AutoStartDisabled = 1,
            SaveState = 2,
            PowerOff = 3,
            AcpiShutdown = 4
        }

        public enum VmState
        {
            Null = 0,
            PoweredOff = 1,
            Saved = 2,
            Aborted = 4,
            Running = 5,
            Paused = 6,
            Stuck = 7,
            Starting = 10,
            Stopping = 11,
            Saving = 12,
            Restoring = 13
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                if (components != null)
                    components.Dispose();

            base.Dispose(disposing);
        }
        private void CatchStatus(ToolStripMenuItem obj, IProgress vmProgress, IMachine vm)
        {
            while (true)
            {
                Thread.Sleep(50);
                string percent = ((int)vm.State > 9) ? $" - {vmProgress.Percent}%" : string.Empty;

                string str = ((VmState)vm.State).ToString();
                obj.Text = $"{obj.Tag} - {str}{percent}";
                obj.Image = SetImg(str);
                obj.GetCurrentParent().Update();

                if (vmProgress.Completed == 1)
                {
                    str = ((VmState)vm.State).ToString();
                    obj.Text = $"{obj.Tag} - {str}";
                    obj.Image = SetImg(str);
                    obj.GetCurrentParent().Update();
                    break;
                }
            }
        }
    }
}