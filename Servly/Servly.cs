using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Management;
using System.Net;
using System.IO;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Security.Permissions;
using WUApiLib;


using Echevil; // networking code used from: http://netcode.ru/dotnet/?lang=&katID=30&skatID=277&artID=7659

namespace Servly
{
    [RunInstaller(true)]
    public partial class ScheduleTask : Installer
    {
        public ScheduleTask()
        {
            //InitializeComponent();
        }

       

        [SecurityPermission(SecurityAction.Demand)]
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);

            RemoveScheduledTask();

            string installationPath = Context.Parameters["sd"] ?? "";
            //Without the replace, results in c:\path\\MyApp.exe
            string apikey = Context.Parameters["apikey"];
            string subdomain = Context.Parameters["subdomain"];
            string servly_url = "http://"+subdomain+".servly.com/status/update?key="+apikey;
        
            string executablePath = Path.Combine(installationPath, "Servly.exe").Replace("\\\\", "\\");

            // http://technet.microsoft.com/en-us/library/cc725744%28WS.10%29.aspx#BKMK_minutes
            Process scheduler = Process.Start("schtasks.exe", string.Format("/Create /RU SYSTEM /SC minute /MO 4 /TN \"Servly\" /TR \"\\\"{0}\\\" "+servly_url+"\"", executablePath));
            scheduler.WaitForExit();
        }

        [SecurityPermission(SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            RemoveScheduledTask();
        }

        private void RemoveScheduledTask()
        {
            Process scheduler = Process.Start("schtasks.exe", "/Delete /TN \"MyApp\" /F");
            scheduler.WaitForExit();
        }
    }

    class Servly
    {
        static void Main(string[] args)
        {
            try
            {
                string servlyUrl = args[0];
                Console.WriteLine(args[0]);
            
                Dictionary<string,string> postData = new Dictionary<string,string>();

                Dictionary<string, double> Disk = DiskUsage();
                Dictionary<string, double> Memory = MemoryUsage();
                Dictionary<string, double> Network = NetworkUsage();
                double cpuUsage = CpuUsage();
                int Procs = ProcsRunning();
                double NetConns = NetworkConnections();

                Console.WriteLine("Processors: {0}", Environment.ProcessorCount);

                Console.WriteLine("Operating System: {0}", Environment.OSVersion.ToString());

                Console.WriteLine("Running Processes: {0}", Procs);

                Console.WriteLine("Memory Usage: " + (Memory["availableMemory"] / Memory["totalMemory"]) * 100 + "%");

                Console.WriteLine("CPU Usage: " + cpuUsage + "%");

                Console.WriteLine("Network Connections: " + NetConns);


                Console.WriteLine("Disk Free/Used/Total: {0} / {1} / {2}", 
                        Disk["free"] / (1024 * 1024 * 1024), 
                        Disk["used"] / (1024 * 1024 * 1024), 
                        (Disk["free"] + Disk["used"]) / (1024 * 1024 * 1024));

                Console.WriteLine("Upload/Download: {0}/{1}", Network["upload"], Network["download"]);

                postData.Add("os", "Windows");
                postData.Add("ps", ProcList());
                postData.Add("cpu_usage", cpuUsage.ToString());
                postData.Add("disk_used", Disk["used"].ToString() );
                postData.Add("disk_size", (Disk["free"] + Disk["used"]).ToString() );
                postData.Add("mem_used", (Memory["totalMemory"] - Memory["availableMemory"] ).ToString() );
                postData.Add("mem_free", Memory["availableMemory"].ToString() );
                postData.Add("running_procs", Procs.ToString() );
                postData.Add("net_in", Network["download"].ToString() );
                postData.Add("net_out", Network["upload"].ToString());
                postData.Add("number_of_cpus", Environment.ProcessorCount.ToString() );
                postData.Add("kernel", Environment.OSVersion.ToString() );
                postData.Add("connections", NetConns.ToString() );
                postData.Add("uptime", GetUptime());
                
                postData.Add("pending_updates", PendingUpdates().ToString());


                // this is what we are sending
                string post_data = Serialize(postData);
                //Console.WriteLine("Post data: " + post_data);
                

                // this is where we will send it
                string uri = servlyUrl;

                // create a request
                HttpWebRequest request = (HttpWebRequest)
                WebRequest.Create(uri); request.KeepAlive = false;
                request.ProtocolVersion = HttpVersion.Version10;
                request.Method = "POST";

                // turn our request string into a byte stream
                byte[] postBytes = Encoding.ASCII.GetBytes(post_data);

                // this is important - make sure you specify type this way
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = postBytes.Length;
                Stream requestStream = request.GetRequestStream();

                // now send it
                requestStream.Write(postBytes, 0, postBytes.Length);
                requestStream.Close();

                // grab te response and print it out to the console along with the status code
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Console.WriteLine(new StreamReader(response.GetResponseStream()).ReadToEnd());
                Console.WriteLine(response.StatusCode);

                //Console.Read();
            }
            catch (Exception e)
            {
                Console.WriteLine("You need to provide the URL. Exception: " + e.ToString());
            }

        }

        private static string Serialize(Dictionary<string,string> input)
        {
            string str = "";
            foreach (KeyValuePair<string, string> k in input)
            {
                str += "srvly[" + k.Key + "]=" + k.Value + "&";
            }
            return str;
        }

        private static Dictionary<string, double> NetworkUsage()
        {
            Dictionary<string, double> temp = new Dictionary<string, double>();

            double DownloadSpeed = 0;
            double UploadSpeed = 0;

            NetworkMonitor monitor = new NetworkMonitor();
            NetworkAdapter[] adapters = monitor.Adapters;

            // Start a timer to obtain new performance counter sample every second.
            monitor.StartMonitoring();

            for (int i = 0; i < 10; i++)
            {
                foreach (NetworkAdapter adapter in adapters)
                {
                    // The DownloadSpeedKbps and UploadSpeedKbps are
                    // double values. 
                   
                    DownloadSpeed += adapter.DownloadSpeedKbps;
                    UploadSpeed += adapter.UploadSpeedKbps;
                }
                System.Threading.Thread.Sleep(1000); // Sleeps for one second.
            }

            // Stop the timer. Properties of adapter become invalid.
            monitor.StopMonitoring();


            temp.Add("upload", (UploadSpeed/8)*1024);
            temp.Add("download", (DownloadSpeed/8)*1024);

            return temp;
        }

        private static string ProcList()
        {
            string proc_list = "";
            ManagementObjectSearcher m = new ManagementObjectSearcher("SELECT * FROM Win32_Process");
            foreach (ManagementObject mm in m.Get()) 
            {
                proc_list += mm["ProcessID"] + "||";
                proc_list += mm["Name"] + "||";
                //Console.WriteLine("-------------------");
                uint pid = (uint)mm["ProcessId"];
                foreach (PropertyData pd in mm.Properties)
                {
                    //Console.WriteLine(pd.Name + ":" + (mm[pd.Name] == null ? "N/A" : mm[pd.Name].ToString()));
                }
                ManagementObjectSearcher m2 = new ManagementObjectSearcher("SELECT * FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess='"+pid.ToString()+"'");
                foreach (ManagementObject mmm in m2.Get())
                {
                    proc_list += mmm["WorkingSet"] + "||";
                    proc_list += mmm["PrivateBytes"] + "||";
                    //Console.WriteLine(mmm["PercentProcessorTime"].GetType().ToString());
                    UInt64 tmpn = (UInt64)mmm["PercentPrivilegedTime"] + (UInt64)mmm["PercentProcessorTime"] + (UInt64)mmm["PercentUserTime"];
                    proc_list += tmpn.ToString() + "||" ;
                    foreach (PropertyData pd in mmm.Properties)
                    {
                        //Console.WriteLine("\t" + pd.Name + ":" + (mmm[pd.Name] == null ? "N/A" : mmm[pd.Name].ToString()));
                    }
                }

                proc_list += "\n";
            }
            

            /*
            string str = "";
            Process[] processlist = Process.GetProcesses();
            foreach(Process p in processlist){
                try
                {
                    str += p.Id + " || * || " + p.PrivateMemorySize64 + "|| " + p.VirtualMemorySize64 + " || " + p.TotalProcessorTime + " || * || " + p.ProcessName + "\n";
                }
                catch (Exception e)
                {
                    str += p.Id + " || * || " + p.PrivateMemorySize64 + "|| " + p.VirtualMemorySize64 + " || * || * || " + p.ProcessName + "\n";
                }
            }

            return str;*/
            return proc_list;
        }

        private static int ProcsRunning()
        {
            Process[] processlist = Process.GetProcesses();
            return processlist.Length;
        }

        private static double CpuUsage()
        {
            // Get CPU usage over 20 seconds
            PerformanceCounter cpuCounter;
            double cpuUsage = 0;
            cpuCounter = new PerformanceCounter();
            cpuCounter.CategoryName = "Processor";
            cpuCounter.CounterName = "% Processor Time";
            cpuCounter.InstanceName = "_Total";
            for (int i = 0; i < 20; i++) { cpuUsage += cpuCounter.NextValue(); Thread.Sleep(1000); }
            cpuUsage /= 20;

            return cpuUsage;
        }

        // source; http://www.tech-archive.net/Archive/DotNet/microsoft.public.dotnet.framework.interop/2007-01/msg00001.html
        private static int PendingUpdates()
        {            
            Type stype = Type.GetTypeFromProgID("Microsoft.Update.Session");
            UpdateSession session = (UpdateSession)Activator.CreateInstance(stype);
            Type itype = Type.GetTypeFromProgID("Microsoft.Update.Installer");
            UpdateInstaller installer = (UpdateInstaller)Activator.CreateInstance(itype);
            IUpdate u;

            IUpdateSearcher us = session.CreateUpdateSearcher();
            ISearchResult sr = us.Search("IsInstalled=0 and Type='Software'");
            UpdateCollection updates = sr.Updates;
            for (int i = 0; i < updates.Count; i++)
            {
                u = updates[i]; //do something with update;show.
                //System.Console.WriteLine(u.Title);
                // System.Console.WriteLine(u.Description);
                //System.Console.WriteLine(" ");
            }
            Console.WriteLine("Updates:{0}", updates.Count);
            
            return updates.Count;
        }

        private static string GetUptime()
        {
            string systemUptime = "";
            ManagementObjectSearcher m = new ManagementObjectSearcher("SELECT * FROM Win32_Operatingsystem");
            foreach (ManagementObject mm in m.Get()) { systemUptime = mm["LastBootUpTime"].ToString(); }
            TimeSpan tmp = (DateTime.Now - ParseCIM(systemUptime));


            return String.Format("{0} days {1} hours {2} minutes {3} seconds", tmp.Days, tmp.Hours, tmp.Minutes, tmp.Seconds ) ;
        }

        private static Dictionary<string,double> MemoryUsage()
        {
            Dictionary<string, double> temp = new Dictionary<string, double>();
            double totalMemory = 0;
            double availableMemory = 0;
            // Get memory usage and information
            ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectSearcher mos2 = new ManagementObjectSearcher("SELECT * FROM Win32_perfRawData_PerfOS_Memory");
            // total memory
            foreach (ManagementObject mo in mos.Get()) { totalMemory = System.Convert.ToDouble(mo["TotalPhysicalMemory"].ToString()); }
            // available memory
            foreach (ManagementObject mo in mos2.Get()) { availableMemory = System.Convert.ToDouble(mo["AvailableBytes"].ToString()); }

            temp.Add("totalMemory", totalMemory);
            temp.Add("availableMemory", availableMemory);

            return temp;
        }

        // source: http://www.dreamincode.net/code/snippet5274.htm
        private static DateTime ParseCIM(string date)
        {
            //datetime object to store the return value
            DateTime parsed = DateTime.MinValue;

            //check date integrity
            if (date != null && date.IndexOf('.') != -1)
            {
                //obtain the date with miliseconds
                string newDate = date.Substring(0, date.IndexOf('.') + 4);

                //check the lenght
                if (newDate.Length == 18)
                {
                    //extract each date component
                    int y = Convert.ToInt32(newDate.Substring(0, 4));
                    int m = Convert.ToInt32(newDate.Substring(4, 2));
                    int d = Convert.ToInt32(newDate.Substring(6, 2));
                    int h = Convert.ToInt32(newDate.Substring(8, 2));
                    int mm = Convert.ToInt32(newDate.Substring(10, 2));
                    int s = Convert.ToInt32(newDate.Substring(12, 2));
                    int ms = Convert.ToInt32(newDate.Substring(15, 3));

                    //compose the new datetime object
                    parsed = new DateTime(y, m, d, h, mm, s, ms);
                }
            }

            //return datetime
            return parsed;
        }

        private static double NetworkConnections()
        {
            // Network Information
            System.Diagnostics.ProcessStartInfo procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c netstat -an");
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            // Do not create the black window.
            procStartInfo.CreateNoWindow = true;
            // Now we create a process, assign its ProcessStartInfo and start it
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo = procStartInfo;
            proc.Start();
            // Get the output into a string
            double result = proc.StandardOutput.ReadToEnd().ToString().Split('\n').Length;
            // Display the command output.
            return result;
        }

        private static Dictionary<string, double> DiskUsage()
        {

            Dictionary<string, double> temp = new Dictionary<string, double>();

            double freeSpace = 0;
            double usedSpace = 0;
            
            // get disk stats
            System.Management.ObjectQuery oQuery = new System.Management.ObjectQuery("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3");
            ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oQuery);
            ManagementObjectCollection oReturnCollection = oSearcher.Get();
            
            //loop through found drives and write out info
            foreach (ManagementObject oReturn in oReturnCollection)
            {
                //Free space in MB
                freeSpace += Convert.ToInt64(oReturn["FreeSpace"]);

                //Used space in MB
                usedSpace += (Convert.ToInt64(oReturn["Size"]) - Convert.ToInt64(oReturn["FreeSpace"]));
            }
            temp.Add("used", usedSpace);
            temp.Add("free", freeSpace);

            return temp;
        }
    }
}