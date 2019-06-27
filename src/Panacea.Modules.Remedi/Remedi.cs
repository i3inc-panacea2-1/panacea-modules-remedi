using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Threading;
using Timer = System.Timers.Timer;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Management;
using Panacea.Modularity.Hardware;
using Panacea.Core;

namespace Panacea.Modules.Remedi
{
    internal class Remedi : IHardwareManager
    {
        #region imports
        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern ulong GetHandSetAPIVersion();

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern ulong DisableBarCodeScan(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool HandSetAPIStart();

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool HandSetAPIStop();

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SetRingLED(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool DisableRing(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool DisableRing2(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SwitchRINGTONEMode(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool CheckHookOnStatus();

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SetChipVolumeLevel(byte left, byte right);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SetMICVolumeLevel(byte mode, byte right);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SetChipMICVolumeLevel(byte mode, byte right);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static extern bool SetMainVolumeLevel(byte mode);

        [DllImport("HandSetAPI", CallingConvention = CallingConvention.Cdecl)]
        private static unsafe extern bool GetMainVolumeLevel(byte* mode);

        public byte MicrophoneVolume { get; set; } = 61;

        public byte HandsetVolume { get; set; } = 66;

        public byte SpeakersVolume { get; set; } = 100;

        public bool UseUsbFix { get; set; } = false;

        #endregion imports

        private byte _ledStatus = 0;
        HardwareStatus _barcodeStatus = HardwareStatus.Off;
        private const byte Off = 0x0;
        private const byte On = 0x1;
        private bool _prevHardwareStatus = true;
        const byte Ringtone0 = 0xf;
        private readonly ILogger _logger;
        private Timer _ScannerTimer;
        private Timer _microphoneTimer;

        internal Remedi(byte handsetSpeakerVolume, byte handsetMicVolume, ILogger logger)
        {
            HandsetVolume = handsetSpeakerVolume;
            MicrophoneVolume = handsetMicVolume;
            _logger = logger;
            _ScannerTimer = new Timer(30000);
            _ScannerTimer.Elapsed += _ScannerTimer_Elapsed;
            _microphoneTimer = new Timer(2000);
            _microphoneTimer.Elapsed += _microphoneTimer_Elapsed;
        }

        private void _microphoneTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            ExecuteSync(() => SetChipMICVolumeLevel(MicrophoneVolume, MicrophoneVolume));
            ExecuteSync(() => SetChipVolumeLevel(HandsetVolume, HandsetVolume));
        }

        static string CurrentPath { get => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }

        static Remedi()
        {
            string path;
            switch (IntPtr.Size)
            {
                case 4:
                    path = "x86";
                    break;
                case 8:
                    path = "x64";
                    break;
                default:
                    throw new Exception("Unsupported architecture");
            }
            LoadLibrary(Path.Combine(
                CurrentPath,
                "HandsetAPI",
                path,
                "HandSetAPI.dll"));
        }

        [DllImport("kernel32.dll")]
        private static extern IntPtr LoadLibrary(string dllToLoad);

        private void _ScannerTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            ExecuteSync(() => {
                DisableBarCodeScan(_barcodeStatus == HardwareStatus.On ? Off : On);
                return true;
            });
        }

        public HardwareStatus LedState
        {
            get { return _ledStatus == 0 ? HardwareStatus.Off : HardwareStatus.On; }
            set
            {
                _ledStatus = value == HardwareStatus.On ? On : Off;
                /*
                ExecuteSync(() =>SetRingLED(_ledStatus));
                ExecuteSync(() => DisableRing(value == HardwareStatus.On ? Off : On));
                */
            }
        }

        public HardwareStatus HandsetState => !_prevHardwareStatus ? HardwareStatus.On : HardwareStatus.Off;

        object _lock = new object();

        object ExecuteSync(ActionDel act)
        {
            lock (_lock)
            {
                return act();
            }
        }

        delegate object ActionDel();

        public bool SpeakersDisabled { get; set; }

        public void Restart()
        {
            ExecuteSync(() =>
            {
                HandSetAPIStop();
                HandSetAPIStart();
                return null;
            });
        }

        class ListenerWindow : NativeWindow
        {
            public event EventHandler<Message> Message;
            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);
                Message?.Invoke(this, m);
            }
        }

        ListenerWindow _window;
        public unsafe void Start()
        {
            _window = new ListenerWindow();
            _window.CreateHandle(new System.Windows.Forms.CreateParams());
            _window.Message += _window_Message;
            UsbNotification.RegisterUsbDeviceNotification(_window.Handle);

            if(!(bool)ExecuteSync(() => HandSetAPIStart()))
            {
                _logger.Error(this, "Unable to initialize Remedi SDK");
            }
            _logger?.Debug("Remedi", ((ulong)ExecuteSync(() => GetHandSetAPIVersion())).ToString("X"));
            ExecuteSync(() => SwitchRINGTONEMode(Ringtone0));
            ExecuteSync(() => DisableRing(On));
            ExecuteSync(() => SetRingLED(Off));
            SetMainVolumeLevel(SpeakersVolume);

            _prevHardwareStatus = (bool)ExecuteSync(() => CheckHookOnStatus());

            //ExecuteSync(() => SetMainVolumeLevel(100));

            ExecuteSync(() => SetChipMICVolumeLevel(MicrophoneVolume, MicrophoneVolume));
            ExecuteSync(() => SetChipVolumeLevel(HandsetVolume, HandsetVolume));
            ExecuteSync(() => DisableBarCodeScan(On));
            _logger?.Debug("Remedi", "Handset API started");
            _ScannerTimer.Start();
            System.Windows.Application.Current.Exit += (oo, ee) =>
            {
                _stop = true;
                //Thread.Sleep(360);
                ExecuteSync(() => HandSetAPIStop());
            };
            _thread = new Thread(() =>
            {
                try
                {
                    while (!_stop)
                    {
                        Thread.Sleep(500);

                        var handsetStatus = (bool)ExecuteSync(() => CheckHookOnStatus());
                        if (handsetStatus == _prevHardwareStatus) continue;
                        _prevHardwareStatus = handsetStatus;
                        OnHandsetStateChanged(!_prevHardwareStatus ? HardwareStatus.On : HardwareStatus.Off);
                    }
                }
                catch
                {
                    //ignore
                }
            })
            {
                IsBackground = true
            };
            _thread.Priority = ThreadPriority.Lowest;
            _thread.Start();
        }
        Thread _thread;
        bool _isRestarting;


        private async void _window_Message(object sender, System.Windows.Forms.Message e)
        {
            Debug.WriteLine(e.Msg);
            if (e.Msg == UsbNotification.WmDevicechange)
            {
                if (_isRestarting) return;
                _isRestarting = true;
                await Task.Delay(5000);
                if (UseUsbFix)
                {
                    var devs = GetRootHubs();
                    foreach (var dev in devs)
                    {
                        Console.WriteLine($"devcon.exe restart *{dev}*");
                        var info = new ProcessStartInfo()
                        {
                            FileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Environment.Is64BitOperatingSystem ? "x64" : "x86", "devcon.exe"),
                            Arguments = $"restart *{dev}*",
                            Verb = "runas",
                            CreateNoWindow = true,
                            UseShellExecute = false
                        };
                        var p = new Process()
                        {
                            StartInfo = info
                        };
                        try
                        {
                            p.Start();
                            p.WaitForExit();
                        }
                        catch { }
                    }
                    await Task.Delay(2000);
                    Restart();
                    await Task.Delay(10000);
                }
                else
                {
                    Restart();
                }
                _isRestarting = false;
            }
        }

        static List<string> GetRootHubs()
        {
            return GetUSBDevices().Where(d => d.DeviceID.Contains("ROOT_HUB")).Select(d => d.DeviceID.Split('\\')[1]).Distinct().ToList();
        }
        static List<USBDeviceInfo> GetUSBDevices()
        {
            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();

            ManagementObjectCollection collection;
            using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))
                collection = searcher.Get();

            foreach (var device in collection)
            {
                devices.Add(new USBDeviceInfo(
                (string)device.GetPropertyValue("DeviceID"),
                (string)device.GetPropertyValue("PNPDeviceID"),
                (string)device.GetPropertyValue("Description"),
                ""
                ));
            }

            collection.Dispose();
            return devices;
        }

        private bool _stop;

        public void StartRinging()
        {

        }

        public void StopRinging()
        {

        }

        public void SetLcdBacklight(HardwareStatus status)
        {

        }

        void OnHandsetStateChanged(HardwareStatus status)
        {
            if (System.Windows.Application.Current != null)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => { HandsetStateChanged?.Invoke(this, status); });
            }
            else
            {
                HandsetStateChanged?.Invoke(this, status);
            }
        }
        public event EventHandler<HardwareStatus> HandsetStateChanged;

        public event EventHandler<HardwareStatus> LcdButtonChange
        {
            add { }
            remove { }
        }

        public event EventHandler<HardwareStatus> PowerButtonChange
        {
            add { }
            remove { }
        }


        public void Stop()
        {
            HandSetAPIStop();
        }


        public HardwareStatus BarcodeScannerState
        {
            get { return _barcodeStatus; }
            set
            {
                _barcodeStatus = value;
                DisableBarCodeScan(value == HardwareStatus.On ? Off : On);
            }
        }

        public void SimulateHandsetState(HardwareStatus status)
        {
            OnHandsetStateChanged(status);
        }

        public void Dispose()
        {

        }
    }

    class USBDeviceInfo
    {
        public USBDeviceInfo(string deviceID, string pnpDeviceID, string description, string path)
        {
            this.DeviceID = deviceID;
            this.PnpDeviceID = pnpDeviceID;
            this.Description = description;
            this.Path = path;
        }
        public string DeviceID { get; private set; }
        public string PnpDeviceID { get; private set; }
        public string Description { get; private set; }

        public string Path { get; set; }
    }
}
