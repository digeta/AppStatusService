using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;

using Microsoft.Win32;

using YemekOtomasyonSvc.Error;
using YemekOtomasyonSvc.Misc;

namespace YemekOtomasyonSvc
{
    public partial class YemekSvc : ServiceBase
    {
        private event ErrorOccured OnError;

        private System.Timers.Timer uptimer;
        private Int32 elapsedUp = 0;

        public YemekSvc()
        {
            InitializeComponent();

            if (!System.Diagnostics.EventLog.SourceExists("YemekOtoKontrol"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "YemekOtoKontrol", "YemekOtoKontrolLog");
            }

            eventLog.Source = "YemekOtoKontrol";
            eventLog.Log = "YemekOtoKontrolLog";
        }

        protected override void OnStart(string[] args)
        {
            eventLog.WriteEntry("Yemekhane otomasyon izleyici servisi başladı");

            this.OnError += Service_OnError;
			
            uptimer = new System.Timers.Timer();
            uptimer.Interval = 1000;
            uptimer.Elapsed += new System.Timers.ElapsedEventHandler(uptimer_Elapsed);
            uptimer.Start();
        }

        protected override void OnStop()
        {
        }
        
        private void NetworkCheck()
        {
            try
            {
                SelectQuery query = new SelectQuery("Win32_NetworkAdapter");
                ManagementObjectSearcher search = new ManagementObjectSearcher(query);

                foreach (ManagementObject result in search.Get())
                {
                    NetworkAdapter adapter = new NetworkAdapter(result);

                    if (adapter.NetConnectionID == "bim" || adapter.NetConnectionID == "turnike")
                    {
                        if (adapter.NetConnectionStatus == 0)
                        {
                            adapter.Enable();
                        }

                        if (adapter.NetConnectionStatus == 1 || adapter.NetConnectionStatus == 3)
                        {
                            adapter.Disable();
                            Thread.Sleep(5000);
                            adapter.Enable();
                        }
                    }
                }


                // First Attempt to Fix
                ManagementObjectSearcher search_2 = new ManagementObjectSearcher(query);

                foreach (ManagementObject result in search_2.Get())
                {
                    NetworkAdapter adapter = new NetworkAdapter(result);
                    
                    if (adapter.NetConnectionID == "bim" || adapter.NetConnectionID == "turnike")
                    {
                        if (adapter.NetConnectionStatus != 2)
                        {
                            eventLog.WriteEntry("Network Error : " + adapter.NetConnectionID + " | " + adapter.NetConnectionStatus + " | " + DateTime.Now);

                            RegistryKey myKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\0015", true);
                            if (myKey != null)
                            {
                                myKey.SetValue("*SpeedDuplex", "0", RegistryValueKind.String);
                                myKey.Close();
                            }
                            
                            adapter.Disable();
                            Thread.Sleep(5000);
                            adapter.Enable();
                        }
                    }
                }


                // Second Attempt to Fix
                ManagementObjectSearcher search_3 = new ManagementObjectSearcher(query);

                foreach (ManagementObject result in search_3.Get())
                {
                    NetworkAdapter adapter = new NetworkAdapter(result);

                    if (adapter.NetConnectionID == "bim" || adapter.NetConnectionID == "turnike")
                    {
                        if (adapter.NetConnectionStatus != 2)
                        {
                            eventLog.WriteEntry("Network Error : " + adapter.NetConnectionID + " | " + adapter.NetConnectionStatus + " | " + DateTime.Now);

                            RegistryKey myKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\0015", true);
                            if (myKey != null)
                            {
                                myKey.SetValue("*SpeedDuplex", "2", RegistryValueKind.String);
                                myKey.Close();
                            }

                            adapter.Disable();
                            Thread.Sleep(5000);
                            adapter.Enable();
                        }
                    }
                }


                // Final Attempt to Fix
                ManagementObjectSearcher search_4 = new ManagementObjectSearcher(query);

                foreach (ManagementObject result in search_4.Get())
                {
                    NetworkAdapter adapter = new NetworkAdapter(result);

                    if (adapter.NetConnectionID == "bim" || adapter.NetConnectionID == "turnike")
                    {
                        if (adapter.NetConnectionStatus != 2)
                        {
                            eventLog.WriteEntry("Network Error : " + adapter.NetConnectionID + " | " + adapter.NetConnectionStatus + " | " + DateTime.Now);

                            RegistryKey myKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\0015", true);
                            if (myKey != null)
                            {
                                myKey.SetValue("*SpeedDuplex", "4", RegistryValueKind.String);
                                myKey.Close();
                            }

                            adapter.Disable();
                            Thread.Sleep(5000);
                            adapter.Enable();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String method = this.GetType().Name + ", " + MethodBase.GetCurrentMethod().Name;
                OnError(this, new ErrorOccuredArgs(method, ex.Message));
            }
            finally
            {
                uptimer.Start();
            }
        }
        
        private void LaunchClient()
        {
            ApplicationLoader.PROCESS_INFORMATION procInfo;
            ApplicationLoader.StartProcessAndBypassUAC(Settings.StartPath, out procInfo);

            if (procInfo.dwProcessId > 0)
            {
                eventLog.WriteEntry("Yemekhane kontrol çalıştırıldı");
            }
            else
            {
                eventLog.WriteEntry("Yemekhane kontrol çalıştırılamadı");
            }
        }

        #region "Events"
        private void uptimer_Elapsed(object sender, EventArgs e)
        {
            try
            {
                elapsedUp++;

                Process[] processes = Process.GetProcessesByName("YemekhaneKontrol");
                if (processes.Length <= 0) LaunchClient();

                if (elapsedUp >= 10)
                {
                    elapsedUp = 0;
                    uptimer.Stop();

                    NetworkCheck();
                }
            }
            catch (Exception ex)
            {
                String method = this.GetType().Name + ", " + MethodBase.GetCurrentMethod().Name;
                OnError(this, new ErrorOccuredArgs(method, ex.Message));
            }
        }
		
        private void Service_OnError(object sender, ErrorOccuredArgs e)
        {
            eventLog.WriteEntry("[Method: " + e.Method + "] --- [Message: " + e.Message + "]");
        }
        
        #endregion
    }

    internal static class NativeMethods
    {

        // Import SetThreadExecutionState Win32 API and necessary flags

        [DllImport("kernel32.dll")]

        public static extern uint SetThreadExecutionState(uint esFlags);

        public const uint ES_CONTINUOUS = 0x80000000;

        public const uint ES_SYSTEM_REQUIRED = 0x00000001;

    }
}
