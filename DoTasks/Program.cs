using System;
using System.Management.Automation;
using System.Threading;

namespace DoTasks
{
    class Program
    {
        private const int TimeWait = 30;
        private const string ProgramName = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE";
        static void Main()
        {
            string script = @"[Cmdletbinding()]
            param(

                [Parameter(mandatory =$false, position = 0)]$waitingTimeInSeconds = 40,
                [Parameter(mandatory =$false, position = 1)]$PathToProgram = ""C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"",
                [Parameter(mandatory =$false, position = 2)] [switch]$WhatIf
                )

#before running the script You'd need to change the Execution policy on the computer to ""unrestricted"" like this:
#Set-ExecutionPolicy unrestricted
#And accepting the change on registry, this MUST BE DONE from a ""powershell Console as Administrator"".

            $getAllAdapters = Get-NetAdapter | where { $_.Status -eq ""up"" -and $_.Name -notlike ""bluetooh*""}
            Write-host ""Disabling All network adapters except bluetooh"" -ForegroundColor Green
            if($WhatIf){
                $getAllAdapters  | Disable-NetAdapter -Confirm:$false -WhatIf
            }
            else{
                $getAllAdapters | Disable-NetAdapter -Confirm:$false
            }


            Write-host ""Network adapters are disabled, Executing Program"" -ForegroundColor Green

#run Exe File
                & $PathToProgram

            Write-host ""Waiting $waitingTimeInSeconds seconds"" -ForegroundColor Red
            Start-Sleep -Seconds $waitingTimeInSeconds

#close the exe file?

#Enable All devices all over again 
            Write-host ""Enabling NetworkAdapters"" -ForegroundColor Green
            if($WhatIf){
                $getAllAdapters |Enable-NetAdapter -whatif
            }
            else{
                $getAllAdapters  |Enable-NetAdapter
            }";

            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                PowerShellInstance.AddScript(script);
                //""Get-NetAdapter | where{$_.status -eq 'Up' -and $_.Name -notlike '*Bluet*'");
                IAsyncResult result = PowerShellInstance.BeginInvoke();
                while (result.IsCompleted == false)
                {
                    Console.WriteLine("Waiting for pipeline to finish...");
                    Thread.Sleep(5000);
                }
                Console.WriteLine("Finished!");
                Console.ReadKey();
            }

            /*

                        //Retrive all network interface using GetAllNetworkInterface() method off NetworkInterface class.

                        NetworkInterface[] niArr = NetworkInterface.GetAllNetworkInterfaces();
                        var test = GetAllAdaptors();

                        //Display all information of NetworkInterface using foreach loop.
                        var re = (from r in niArr
                                  where ((!r.Name.Contains("Bluetooh")) && (r.OperationalStatus == OperationalStatus.Up))
                                  select r).ToArray();


                        foreach (NetworkInterface ni in re)
                        {
                            //Console.WriteLine(tempNetworkInterface.Name);
                            Disable(ni.Name);
                            Thread.Sleep(1 * 1000);
                        }


                        StartProgram(ProgramName);



                        Thread.Sleep(TimeWait * 1000);



                        foreach (NetworkInterface ni in re)
                        {
                            Enable(ni.Name);
                            Thread.Sleep(1 * 1000);
                        }

                    }


                    public static IEnumerable<ManagementObject> GetAllAdaptors()
                    {


                        var tempList = new List<ManagementObject>();

                        ManagementObjectCollection moc =
                            new ManagementClass("Win32_NetworkAdapterConfiguration").GetInstances();

                        /*foreach (ManagementObject obj in moc)
                        {
                            tempList.Add(obj);
                        }

                        return tempList;
                          */
        }

        static void Enable(string interfaceName)
        {
            System.Diagnostics.ProcessStartInfo psi =
                new System.Diagnostics.ProcessStartInfo("netsh", "interface set interface \"" + interfaceName + "\" enable");
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo = psi;
            p.Start();
        }

        static void StartProgram(string pname)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(pname);
            System.Diagnostics.Process p = new System.Diagnostics.Process();

            p.StartInfo = psi;
            p.Start();
        }

        static void Disable(string interfaceName)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("netsh", "interface set interface \"" + interfaceName + "\" disable");
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo = psi;
            p.Start();
        }
    }
}


/*
            Console.WriteLine("Network Discription  :  " + tempNetworkInterface.Description);
            Console.WriteLine("Network ID  :  " + tempNetworkInterface.Id);
            Console.WriteLine("Network Name  :  " + tempNetworkInterface.Name);
            Console.WriteLine("Network interface type  :  " + tempNetworkInterface.NetworkInterfaceType.ToString());
            Console.WriteLine("Network Operational Status   :   " + tempNetworkInterface.OperationalStatus.ToString());
            Console.WriteLine("Network Spped   :   " + tempNetworkInterface.Speed);
            Console.WriteLine("Support Multicast   :   " + tempNetworkInterface.SupportsMulticast);
            Console.WriteLine();
            */
