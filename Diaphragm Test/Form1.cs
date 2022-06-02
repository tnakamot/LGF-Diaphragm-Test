/*************Revision D 2.0****************/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;


namespace Diaphragm_Test
{
    public partial class Form1 : Form
    {
        //***************************USB******************************
        internal const uint DIGCF_PRESENT = 0x02;
        internal const uint DIGCF_DEVICEINTERFACE = 0x10;
        //Constants for CreateFile() and other file I/O functions
        internal const short FILE_ATTRIBUTE_NORMAL = 0x80;
        internal const short INVALID_HANDLE_VALUE = -1;
        internal const uint GENERIC_READ = 0x80000000;
        internal const uint GENERIC_WRITE = 0x40000000;
        internal const uint CREATE_NEW = 1;
        internal const uint CREATE_ALWAYS = 2;
        internal const uint OPEN_EXISTING = 3;
        internal const uint FILE_SHARE_READ = 0x00000001;
        internal const uint FILE_SHARE_WRITE = 0x00000002;
        //Constant definitions for certain WM_DEVICECHANGE messages
        internal const uint WM_DEVICECHANGE = 0x0219;
        internal const uint DBT_DEVICEARRIVAL = 0x8000;
        internal const uint DBT_DEVICEREMOVEPENDING = 0x8003;
        internal const uint DBT_DEVICEREMOVECOMPLETE = 0x8004;
        internal const uint DBT_CONFIGCHANGED = 0x0018;
        //Other constant definitions
        internal const uint DBT_DEVTYP_DEVICEINTERFACE = 0x05;
        internal const uint DEVICE_NOTIFY_WINDOW_HANDLE = 0x00;
        internal const uint ERROR_SUCCESS = 0x00;
        internal const uint ERROR_NO_MORE_ITEMS = 0x00000103;
        internal const uint SPDRP_HARDWAREID = 0x00000001;

        internal const String reportPath = "c:\\Test Reports\\";

        internal struct SP_DEVICE_INTERFACE_DATA
        {
            internal uint cbSize;               //DWORD
            internal Guid InterfaceClassGuid;   //GUID
            internal uint Flags;                //DWORD
            internal uint Reserved;             //ULONG_PTR MSDN says ULONG_PTR is "typedef unsigned __int3264 ULONG_PTR;"  
        }

        internal struct SP_DEVICE_INTERFACE_DETAIL_DATA
        {
            internal uint cbSize;               //DWORD
            internal char[] DevicePath;         //TCHAR array of any size
        }

        internal struct SP_DEVINFO_DATA
        {
            internal uint cbSize;       //DWORD
            internal Guid ClassGuid;    //GUID
            internal uint DevInst;      //DWORD
            internal uint Reserved;     //ULONG_PTR  MSDN says ULONG_PTR is "typedef unsigned __int3264 ULONG_PTR;"  
        }

        internal struct DEV_BROADCAST_DEVICEINTERFACE
        {
            internal uint dbcc_size;            //DWORD
            internal uint dbcc_devicetype;      //DWORD
            internal uint dbcc_reserved;        //DWORD
            internal Guid dbcc_classguid;       //GUID
            internal char[] dbcc_name;          //TCHAR array
        }
        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern IntPtr SetupDiGetClassDevs(
            ref Guid ClassGuid,     //LPGUID    Input: Need to supply the class GUID. 
            IntPtr Enumerator,      //PCTSTR    Input: Use NULL here, not important for our purposes
            IntPtr hwndParent,      //HWND      Input: Use NULL here, not important for our purposes
            uint Flags);            //DWORD     Input: Flags describing what kind of filtering to use.
        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiEnumDeviceInterfaces(
            IntPtr DeviceInfoSet,           //Input: Give it the HDEVINFO we got from SetupDiGetClassDevs()
            IntPtr DeviceInfoData,          //Input (optional)
            ref Guid InterfaceClassGuid,    //Input 
            uint MemberIndex,               //Input: "Index" of the device you are interested in getting the path for.
            ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData);    //Output: This function fills in an "SP_DEVICE_INTERFACE_DATA" structure.

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiDestroyDeviceInfoList(
            IntPtr DeviceInfoSet);          //Input: Give it a handle to a device info list to deallocate from RAM.

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiEnumDeviceInfo(
            IntPtr DeviceInfoSet,
            uint MemberIndex,
            ref SP_DEVINFO_DATA DeviceInterfaceData);

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiGetDeviceRegistryProperty(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            uint Property,
            ref uint PropertyRegDataType,
            IntPtr PropertyBuffer,
            uint PropertyBufferSize,
            ref uint RequiredSize);

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiGetDeviceInterfaceDetail(
            IntPtr DeviceInfoSet,                   //Input: Wants HDEVINFO which can be obtained from SetupDiGetClassDevs()
            ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData,                    //Input: Pointer to an structure which defines the device interface.  
            IntPtr DeviceInterfaceDetailData,      //Output: Pointer to a SP_DEVICE_INTERFACE_DETAIL_DATA structure, which will receive the device path.
            uint DeviceInterfaceDetailDataSize,     //Input: Number of bytes to retrieve.
            ref uint RequiredSize,                  //Output (optional): The number of bytes needed to hold the entire struct 
            IntPtr DeviceInfoData);                 //Output (optional): Pointer to a SP_DEVINFO_DATA structure

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SetupDiGetDeviceInterfaceDetail(
            IntPtr DeviceInfoSet,                   //Input: Wants HDEVINFO which can be obtained from SetupDiGetClassDevs()
            ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData,               //Input: Pointer to an structure which defines the device interface.  
            IntPtr DeviceInterfaceDetailData,       //Output: Pointer to a SP_DEVICE_INTERFACE_DETAIL_DATA structure, which will contain the device path.
            uint DeviceInterfaceDetailDataSize,     //Input: Number of bytes to retrieve.
            IntPtr RequiredSize,                    //Output (optional): Pointer to a DWORD to tell you the number of bytes needed to hold the entire struct 
            IntPtr DeviceInfoData);                 //Output (optional): Pointer to a SP_DEVINFO_DATA structure

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern IntPtr RegisterDeviceNotification(
            IntPtr hRecipient,
            IntPtr NotificationFilter,
            uint Flags);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern SafeFileHandle CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool WriteFile(
            SafeFileHandle hFile,
            byte[] lpBuffer,
            uint nNumberOfBytesToWrite,
            ref uint lpNumberOfBytesWritten,
            IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool ReadFile(
            SafeFileHandle hFile,
            IntPtr lpBuffer,
            uint nNumberOfBytesToRead,
            ref uint lpNumberOfBytesRead,
            IntPtr lpOverlapped);
        //--------------- Global Varibles Section ------------------
        bool AttachedState = false;                     //Need to keep track of the USB device attachment status for proper plug and play operation.
        bool AttachedButBroken = false;
        SafeFileHandle WriteHandleToUSBDevice = null;
        SafeFileHandle ReadHandleToUSBDevice = null;
        String DevicePath = null;   //Need the find the proper device path before you can open file handles.

        Guid InterfaceClassGuid = new Guid(0x4d1e55b2, 0xf16f, 0x11cf, 0x88, 0xcb, 0x00, 0x11, 0x11, 0x00, 0x00, 0x30);
        //---------------------------------------------------------
        bool OfflineMode = false;
        byte[] rxBuffer = new byte[4096];
        byte[] StreamBuffer = new byte[21];
        long RawForce = 0;
        long RawPosition = 0;
        long RawCurrent = 0;
        long MaxRawCurrent = 0;
        short RampFreq = 0;
        double Speed = 0;
        double FSpeed = 0;
        double SlopeRate = 0;
        double Force = 0;
        double OldForce = 0;
        double Stroke = 0;
        double OldStroke = 0;
        double OldTime = 0;
        bool GetDiff = false;
        int ControlWord = 0;
        long PR2 = 0;
        int StatusWord = 0;
        double MaxCurrent = 0;
        double MaxForce = 0;
        double MaxStroke = 0;
        int CycleNumber = 0;
        int CycleSP = 0;
        int State = 0;
        int OldState = 0;
        double Current = 0;
        bool Auto = true;
        bool StartTest = false;
        long LoadOffset = 0;
        double LoadGain = 1;
        String[] lines = new String[11];
        long TimeCount = 0;
        double Time = 0;
        double[] ForceData = new double[8640000];
        double[] StrokeData = new double[8640000];
        double[] CurrentData = new double[8640000];
        double[] TimeData = new double[8640000];
        int[] IndexData = new int[8640000];
        int Index = 0;
        long DataCount = 0;
        bool CollectData = false;
        string TimeStamp;
        string Date;
        string StatusText = " ";
        double StartTime = 0;
        double TimeRemaining = 0;
        bool DataReceived = false;
        bool StatusReceived = false;
        double RampTimeMultiply = 0;
        string ConfigFilePath = Application.StartupPath + @"\Config.txt";
        string Password = "1234";
        double OldCurrent = 0;
        double SpeedRatio = 13.2;
        double CurrentGain = 1.003342086416806;
        long OldRawTime = -10;
        bool SendOnce = false;
        bool SendCofig = false;
        bool SendControl = false;
        int AnalogSP = 0;
        bool WaitToConnect = true;
        bool PowerError = false;
        bool TempError = false;
        bool ForceLimit = false;
        int SpeedTimer = 0;
        bool LoadDataOnce = false;
        int TestIndex = 0;
        int StartDelay = 0;
        bool StartDelayDone = false;
        double ScaleFactor = 0;
        double PositionOffset = 0;
        //--------------- End of Global Varibles ------------------
        public Form1()
        {
            InitializeComponent();
            DEV_BROADCAST_DEVICEINTERFACE DeviceBroadcastHeader = new DEV_BROADCAST_DEVICEINTERFACE();
            DeviceBroadcastHeader.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
            DeviceBroadcastHeader.dbcc_size = (uint)Marshal.SizeOf(DeviceBroadcastHeader);
            DeviceBroadcastHeader.dbcc_reserved = 0;	
            DeviceBroadcastHeader.dbcc_classguid = InterfaceClassGuid;

            IntPtr pDeviceBroadcastHeader = IntPtr.Zero;  
            pDeviceBroadcastHeader = Marshal.AllocHGlobal(Marshal.SizeOf(DeviceBroadcastHeader)); 
            Marshal.StructureToPtr(DeviceBroadcastHeader, pDeviceBroadcastHeader, false); 
            RegisterDeviceNotification(this.Handle, pDeviceBroadcastHeader, DEVICE_NOTIFY_WINDOW_HANDLE);


            if (CheckIfPresentAndGetUSBDevicePath())    
            {
                uint ErrorStatusWrite;
                uint ErrorStatusRead;


                WriteHandleToUSBDevice = CreateFile(DevicePath, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                ErrorStatusWrite = (uint)Marshal.GetLastWin32Error();
                ReadHandleToUSBDevice = CreateFile(DevicePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                ErrorStatusRead = (uint)Marshal.GetLastWin32Error();

                if ((ErrorStatusWrite == ERROR_SUCCESS) && (ErrorStatusRead == ERROR_SUCCESS))
                {
                    AttachedState = true;       
                    AttachedButBroken = false;
                }
                else 
                {
                    AttachedState = false;      
                    AttachedButBroken = true;   
                    if (ErrorStatusWrite == ERROR_SUCCESS)
                        WriteHandleToUSBDevice.Close();
                    if (ErrorStatusRead == ERROR_SUCCESS)
                        ReadHandleToUSBDevice.Close();
                }
            }
            else    
            {
                AttachedState = false;
                AttachedButBroken = false;
            }

            if (AttachedState == true)
            {
                DeviceStatusLabel.Text = "Device Found";
            }
            else
            {
                DeviceStatusLabel.Text = "Device not found";
            }

            ReadWriteThread.RunWorkerAsync();   

            PasswordTextBox.PasswordChar = '*';

            try
            {
                tabPage1.Controls.Add(ForceStrokeChart);
                ForceStrokeChart.SetBounds(0, 0, 1360, 940);

                tabPage2.Controls.Add(CurrentStrokeChart);
                CurrentStrokeChart.SetBounds(0, 0, 1360, 940);

                tabPage3.Controls.Add(ForceTimeChart);
                ForceTimeChart.SetBounds(0, 0, 1360, 940);
/*
                tabPage4.Controls.Add(ForceStrokeChart);
                ForceTimeChart.SetBounds(0, 0, 1360, 310);
                tabPage4.Controls.Add(CurrentStrokeChart);
                ForceTimeChart.SetBounds(0, 315, 1360, 310);
                tabPage4.Controls.Add(ForceTimeChart);
                ForceTimeChart.SetBounds(0, 630, 1360, 310);
*/

            } catch (Exception e) { Console.WriteLine(e); }
        }
        bool CheckIfPresentAndGetUSBDevicePath()
        {

            try
            {
                IntPtr DeviceInfoTable = IntPtr.Zero;
                SP_DEVICE_INTERFACE_DATA InterfaceDataStructure = new SP_DEVICE_INTERFACE_DATA();
                SP_DEVICE_INTERFACE_DETAIL_DATA DetailedInterfaceDataStructure = new SP_DEVICE_INTERFACE_DETAIL_DATA();
                SP_DEVINFO_DATA DevInfoData = new SP_DEVINFO_DATA();

                uint InterfaceIndex = 0;
                uint dwRegType = 0;
                uint dwRegSize = 0;
                uint dwRegSize2 = 0;
                uint StructureSize = 0;
                IntPtr PropertyValueBuffer = IntPtr.Zero;
                bool MatchFound = false;
                uint ErrorStatus;
                uint LoopCounter = 0;

                String DeviceIDToFind = "Vid_04d8&Pid_003f";

                DeviceInfoTable = SetupDiGetClassDevs(ref InterfaceClassGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

                if (DeviceInfoTable != IntPtr.Zero)
                {
                    while (true)
                    {
                        InterfaceDataStructure.cbSize = (uint)Marshal.SizeOf(InterfaceDataStructure);
                        if (SetupDiEnumDeviceInterfaces(DeviceInfoTable, IntPtr.Zero, ref InterfaceClassGuid, InterfaceIndex, ref InterfaceDataStructure))
                        {
                            ErrorStatus = (uint)Marshal.GetLastWin32Error();
                            if (ErrorStatus == ERROR_NO_MORE_ITEMS) 
                            {   
                                SetupDiDestroyDeviceInfoList(DeviceInfoTable);  
                                return false;
                            }
                        }
                        else    
                        {
                            ErrorStatus = (uint)Marshal.GetLastWin32Error();
                            SetupDiDestroyDeviceInfoList(DeviceInfoTable);  
                            return false;
                        }
                        DevInfoData.cbSize = (uint)Marshal.SizeOf(DevInfoData);
                        SetupDiEnumDeviceInfo(DeviceInfoTable, InterfaceIndex, ref DevInfoData);

                        SetupDiGetDeviceRegistryProperty(DeviceInfoTable, ref DevInfoData, SPDRP_HARDWAREID, ref dwRegType, IntPtr.Zero, 0, ref dwRegSize);
                        PropertyValueBuffer = Marshal.AllocHGlobal((int)dwRegSize);
                        SetupDiGetDeviceRegistryProperty(DeviceInfoTable, ref DevInfoData, SPDRP_HARDWAREID, ref dwRegType, PropertyValueBuffer, dwRegSize, ref dwRegSize2);
                        String DeviceIDFromRegistry = Marshal.PtrToStringUni(PropertyValueBuffer); 

                        Marshal.FreeHGlobal(PropertyValueBuffer);      
                        DeviceIDFromRegistry = DeviceIDFromRegistry.ToLowerInvariant();
                        DeviceIDToFind = DeviceIDToFind.ToLowerInvariant();
                        MatchFound = DeviceIDFromRegistry.Contains(DeviceIDToFind);
                        if (MatchFound == true)
                        {
                            DetailedInterfaceDataStructure.cbSize = (uint)Marshal.SizeOf(DetailedInterfaceDataStructure);
                            SetupDiGetDeviceInterfaceDetail(DeviceInfoTable, ref InterfaceDataStructure, IntPtr.Zero, 0, ref StructureSize, IntPtr.Zero);
                            IntPtr pUnmanagedDetailedInterfaceDataStructure = IntPtr.Zero; 
                            pUnmanagedDetailedInterfaceDataStructure = Marshal.AllocHGlobal((int)StructureSize);    
                            DetailedInterfaceDataStructure.cbSize = 6; 
                            Marshal.StructureToPtr(DetailedInterfaceDataStructure, pUnmanagedDetailedInterfaceDataStructure, false); 
                            if (SetupDiGetDeviceInterfaceDetail(DeviceInfoTable, ref InterfaceDataStructure, pUnmanagedDetailedInterfaceDataStructure, StructureSize, IntPtr.Zero, IntPtr.Zero))
                            {
                                IntPtr pToDevicePath = new IntPtr((uint)pUnmanagedDetailedInterfaceDataStructure.ToInt32() + 4);  
                                DevicePath = Marshal.PtrToStringUni(pToDevicePath); 
                                SetupDiDestroyDeviceInfoList(DeviceInfoTable);	
                                Marshal.FreeHGlobal(pUnmanagedDetailedInterfaceDataStructure);  
                                return true;
                            }
                            else 
                            {
                                uint ErrorCode = (uint)Marshal.GetLastWin32Error();
                                SetupDiDestroyDeviceInfoList(DeviceInfoTable);	
                                Marshal.FreeHGlobal(pUnmanagedDetailedInterfaceDataStructure);
                                return false;
                            }
                        }

                        InterfaceIndex++;
                        LoopCounter++;
                        if (LoopCounter == 10000000)    
                        {
                            return false;
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_DEVICECHANGE)
            {
                if (((int)m.WParam == DBT_DEVICEARRIVAL) || ((int)m.WParam == DBT_DEVICEREMOVEPENDING) || ((int)m.WParam == DBT_DEVICEREMOVECOMPLETE) || ((int)m.WParam == DBT_CONFIGCHANGED))
                {
                    if (CheckIfPresentAndGetUSBDevicePath())	
                    {
                        if ((AttachedState == false) || (AttachedButBroken == true))	
                        {
                            uint ErrorStatusWrite;
                            uint ErrorStatusRead;
                            WriteHandleToUSBDevice = CreateFile(DevicePath, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                            ErrorStatusWrite = (uint)Marshal.GetLastWin32Error();
                            ReadHandleToUSBDevice = CreateFile(DevicePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                            ErrorStatusRead = (uint)Marshal.GetLastWin32Error();

                            if ((ErrorStatusWrite == ERROR_SUCCESS) && (ErrorStatusRead == ERROR_SUCCESS))
                            {
                                AttachedState = true;		
                                AttachedButBroken = false;
                                DeviceStatusLabel.Text = "Device Found, AttachedState = TRUE";
                            }
                            else 
                            {
                                AttachedState = false;		
                                AttachedButBroken = true;	
                                if (ErrorStatusWrite == ERROR_SUCCESS)
                                    WriteHandleToUSBDevice.Close();
                                if (ErrorStatusRead == ERROR_SUCCESS)
                                    ReadHandleToUSBDevice.Close();
                            }
                        }
                    }
                    else	
                    {
                        if (AttachedState == true)		
                        {
                            AttachedState = false;
                            WriteHandleToUSBDevice.Close();
                            ReadHandleToUSBDevice.Close();
                        }
                        AttachedState = false;
                        AttachedButBroken = false;
                    }
                }
            } 

            base.WndProc(ref m);
        } //end of: WndProc() function
        private void ReadWriteThread_DoWork(object sender, DoWorkEventArgs e)
        {
            Byte[] OUTBuffer = new byte[65];    
            Byte[] INBuffer = new byte[65];     
            uint BytesWritten = 0;
            uint BytesRead = 0;

            while (true)
            {
                try
                {
                    if (AttachedState == true)	
                    {
                        if (SendControl == true)
                        {

                            OUTBuffer[0] = 0;	
                            OUTBuffer[1] = 0x81;
                            OUTBuffer[2] = (byte)(ControlWord);
                            OUTBuffer[3] = (byte)(ControlWord >> 8);
                            for (uint i = 4; i < 65; i++)	
                                OUTBuffer[i] = 0xFF;		
                            WriteFile(WriteHandleToUSBDevice, OUTBuffer, 65, ref BytesWritten, IntPtr.Zero);	
                            SendControl = false;
                            if (ForceLimit)
                            {
                              ControlWord = 1;
                              SendControl = true;
                              ForceLimit = false;
                            }
                        }

                        OUTBuffer[0] = 0x00;	
                        OUTBuffer[1] = 0x37;	
                        for (uint i = 2; i < 65; i++)
                            OUTBuffer[i] = 0xFF;

                        if (WriteFile(WriteHandleToUSBDevice, OUTBuffer, 65, ref BytesWritten, IntPtr.Zero))
                        {
                            INBuffer[0] = 0;
                            if (ReadFileManagedBuffer(ReadHandleToUSBDevice, INBuffer, 65, ref BytesRead, IntPtr.Zero))		
                            {
                                if (INBuffer[1] == 0x37)
                                {
                                    RawForce = (INBuffer[2] + INBuffer[3] * 0x100 + INBuffer[4] * 0x10000) & 0xFFFFE0;
                                    RawPosition = (INBuffer[9] * 0x1000000) + (INBuffer[8] * 0x10000) + (INBuffer[7] * 0x100) + INBuffer[6];
                                    RawCurrent = ((INBuffer[12] & 0x7f) << 16) + (INBuffer[11] << 8) + INBuffer[10];
                                    TimeCount = (uint)(INBuffer[17] << 24) + (uint)(INBuffer[16] << 16) + (uint)(INBuffer[15] << 8) + INBuffer[14];
                                    Index = INBuffer[18];
                                    StatusWord = (INBuffer[20] * 0x100) + INBuffer[19];
                                    State = INBuffer[21];
                                    if (State > OldState )
                                    {
                                        if (State == 10 || State == 30 || State == 50 || State == 70)
                                            if (!SendOnce) SendOnce = true;
                                        StartDelayDone = false;
                                        StartDelay = 0; 
                                    }
                                    OldState = State;
                                    if (RawCurrent >= 0x20000)
                                        RawCurrent -= 0x20000;
                                    else
                                        RawCurrent += 0x20000;
                                    if ((INBuffer[12] & 0x80) == 0x80)
                                    {
                                        Current = ((double)(RawCurrent) / 262144 * -10000) * CurrentGain;
                                    }
                                    else
                                        Current = ((double)(RawCurrent) / 262144 * 10000) * CurrentGain;

                                    if (StatusWord > 0 && !StatusReceived) StatusReceived = true;
                                    Time = TimeCount * 5;
                                    Stroke = (double)((RawPosition * ScaleFactor));
                                    Force = Math.Round((double)(RawForce - LoadOffset) * LoadGain / 18480, 2);
                                    DataReceived = true;
                                    if (TimeCount == 0)
                                    {
                                        TimeCount++;
                                    }
                                    if (Index == 1)
                                    {
                                        TestIndex++;
                                        IndexData[DataCount-1] = Index;
                                    }

                                    if (TimeCount > OldRawTime)
                                    {
                                        StrokeData[DataCount] = Math.Round(Stroke,0);
                                        ForceData[DataCount] = Force;
                                        CurrentData[DataCount] = Current;
                                        TimeData[DataCount] = TimeCount*5;
                                        DataCount++;
                                    }
                                    OldRawTime = TimeCount;
                                }
                            }
                        }
                        if (SendCofig == true)
                        {

                            OUTBuffer[0] = 0;				
                            OUTBuffer[1] = 0x80;            

                            OUTBuffer[2] = (byte)(AnalogSP & 0x000000FF);
                            OUTBuffer[3] = (byte)((AnalogSP & 0x0000FF00) >> 8);
                            OUTBuffer[4] = (byte)((AnalogSP & 0x00FF0000) >> 16);
                            OUTBuffer[5] = (byte)((AnalogSP & 0xFF000000) >> 24);

                            OUTBuffer[6] = (byte)(PR2 & 0x000000FF);
                            OUTBuffer[7] = (byte)((PR2 & 0x0000FF00) >> 8);
                            OUTBuffer[8] = (byte)((PR2 & 0x00FF0000) >> 16);
                            OUTBuffer[9] = (byte)((PR2 & 0xFF000000) >> 24);

                            OUTBuffer[10] = (byte)(MaxRawCurrent & 0x000000FF);
                            OUTBuffer[11] = (byte)((MaxRawCurrent & 0x0000FF00) >> 8);
                            OUTBuffer[12] = (byte)((MaxRawCurrent & 0x00FF0000) >> 16);
                            OUTBuffer[13] = 0;

                            OUTBuffer[14] = (byte)((int)MaxStroke & 0x000000FF);
                            OUTBuffer[15] = (byte)(((int)MaxStroke & 0x0000FF00) >> 8);
                            OUTBuffer[16] = (byte)((int)PositionOffset & 0x000000FF);
                            OUTBuffer[17] = (byte)(((int)PositionOffset & 0x0000FF00) >> 8);

                            OUTBuffer[18] = (byte)(((int)(ScaleFactor * 1000) & 0x000000FF));
                            OUTBuffer[19] = (byte)(((int)(ScaleFactor * 1000) & 0x0000FF00) >> 8);


                            for (uint i = 20; i < 65; i++)	
                                OUTBuffer[i] = 0xFF;		
                            WriteFile(WriteHandleToUSBDevice, OUTBuffer, 65, ref BytesWritten, IntPtr.Zero);	
                            SendCofig = false;
                        }
                    } 
                    else
                    {
                        Thread.Sleep(5);    
                                            
                    }
                }
                catch
                {
                }

            } 
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------
        }
        public unsafe bool ReadFileManagedBuffer(SafeFileHandle hFile, byte[] INBuffer, uint nNumberOfBytesToRead, ref uint lpNumberOfBytesRead, IntPtr lpOverlapped)
        {
            IntPtr pINBuffer = IntPtr.Zero;

            try
            {
                pINBuffer = Marshal.AllocHGlobal((int)nNumberOfBytesToRead);    

                if (ReadFile(hFile, pINBuffer, nNumberOfBytesToRead, ref lpNumberOfBytesRead, lpOverlapped))
                {
                    Marshal.Copy(pINBuffer, INBuffer, 0, (int)lpNumberOfBytesRead);
                    Marshal.FreeHGlobal(pINBuffer);
                    return true;
                }
                else
                {
                    Marshal.FreeHGlobal(pINBuffer);
                    return false;
                }

            }
            catch
            {
                if (pINBuffer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pINBuffer);
                }
                return false;
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!OfflineMode)
            {
                if (AttachedState == true && !PowerError && !TempError)
                {
                    DeviceStatusLabel.Text = "Device Found";
                    EnableControls();
                    if (WaitToConnect)
                    {
                        WaitToConnect = false;
                        LoadData();
                    }
                }
                if ((AttachedState == false) || (AttachedButBroken == true) || PowerError || TempError)
                {
                    DeviceStatusLabel.Text = "Device Not Found";
                    StatusLabel.Text = "Device Not Connected";
                    TempStatusLabel.Text = "Temperature State Unknown";
                    PowerStatusLabel.Text = "Power State Unknown";
                    DisableControls();
                }
                if (AttachedState == true)
                {
                    if (CollectData && DataReceived)
                    {
                        this.ForceStrokeChart.Series[0].Points.AddXY(Math.Round(Stroke, 0), Force);
                        if (this.ForceStrokeChart.Series[0].Points.Count > 2)
                        {
                            this.ForceStrokeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count - 1].MarkerSize = 6;
                            this.ForceStrokeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count - 2].MarkerSize = 0;
                            this.ForceStrokeChart.Series[0].Points[0].MarkerSize = 0;
                        }

                        this.CurrentStrokeChart.Series[0].Points.AddXY(Math.Round(Stroke, 0), Current);
                        if (this.CurrentStrokeChart.Series[0].Points.Count > 2)
                        {
                            this.CurrentStrokeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count-1].MarkerSize = 6;
                            this.CurrentStrokeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count-2].MarkerSize = 0;
                            this.CurrentStrokeChart.Series[0].Points[0].MarkerSize = 0;
                        }
                        this.ForceTimeChart.Series[0].Points.AddXY(Math.Round(Time, 2), Force);
                        if (this.ForceTimeChart.Series[0].Points.Count > 2)
                        {
                            this.ForceTimeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count - 1].MarkerSize = 6;
                            this.ForceTimeChart.Series[0].Points[this.CurrentStrokeChart.Series[0].Points.Count - 2].MarkerSize = 0;
                            this.ForceTimeChart.Series[0].Points[0].MarkerSize = 0;
                        }

                        if (Math.Abs(Force) > MaxForce && SendOnce)
                        {
                            ControlWord = 0x101;
                            SendControl = true;
                            ForceLimit = true;
                            SendOnce = false;
                        }
                        DataReceived = false;
                    }
                    OldCurrent = Current;
                    DisplayValues();
                    if ((StatusWord & 1) == 1) Auto = true; else Auto = false;
                    switch (State)
                    {
                        case 0:
                            StatusText = "Stopped";
                            break;
                        case 10:
                            StatusText = "First Cycle:Moving Positive Towards Maximum Force";
                            break;
                        case 20:
                            StatusText = "First Cycle:Moving Negative Towards Zero Force";
                            break;
                        case 30:
                            StatusText = "First Cycle:Moving Negative Towards Maximum Force";
                            break;
                        case 40:
                            StatusText = "First Cycle:Moving Positive Towards Zero Force";
                            break;
                        case 50:
                            StatusText = "Second Cycle:Moving Positive Towards Maximum Force";
                            break;
                        case 60:
                            StatusText = "Second Cycle:Moving Negative Towards Zero Force";
                            break;
                        case 70:
                            StatusText = "Second Cycle:Moving Negative Towards Maximum Force";
                            break;
                        case 80:
                            StatusText = "Second Cycle:Moving Positive Towards Zero Force";
                            break;

                    }
//                    if ((StatusWord & 32) == 32 && State == 0)
                    if ((StatusWord & 32) == 32 && State == 0)
                        {
                            if (CycleNumber < CycleSP)
                        {
                            timer1.Enabled = false;
                            StatusLabel.Text = "Saving";
                            SaveData(CycleNumber);
                            CycleNumber++;
                            timer1.Enabled = true;
                            Restart();
                        }
                        else
                        {
                            ProgressBar.Value = ProgressBar.Maximum;
                            Stop();
                        }
                    }
                    if (State > 0 && DataReceived)
                    {
                        if ((Time / 1000) < ProgressBar.Maximum) ProgressBar.Value = (int)Time / 1000;
                        TimeRemaining = ProgressBar.Maximum - (Time / 1000);
                        TimeSpan t = TimeSpan.FromSeconds(TimeRemaining);
                        TimeRemainingTextBox.Text = t.Minutes.ToString("00") + ":" + t.Seconds.ToString("00");
                    }
                }

            }
            else
            {
                StatusLabel.Text = "Offline Mode";
                EnableControls();
                if (LoadDataOnce)
                {
                    LoadData();
                    LoadDataOnce = false;
                }
                DisplayValues();


            }
        }
        void EnableControls()
        {
            if (LogGroupBox.Enabled == false)
            {
                LogGroupBox.Enabled = true;
                ControlGroupBox.Enabled = true;
                ConfigGroupBox.Enabled = true;
                ChartControlGroupBox.Enabled = true;
            }

        }
        void DisableControls()
        {
            LogGroupBox.Enabled = false;
            ControlGroupBox.Enabled = false;
            ConfigGroupBox.Enabled = false;
            ChartControlGroupBox.Enabled = false;
        }
        double GetTime()
        {
            string TimeString = DateTime.Now.ToString("hh:mm:ss").ToString();
            TimeSpan ts = TimeSpan.Parse(TimeString);
            double second = ts.TotalSeconds;

            return second;
        }
        void DisplayValues()
        {
            if (!OfflineMode)
            {
                if (State>0 && !StartDelayDone)
                {
                    StartDelay++;
                    if (StartDelay>200)
                    {
                        StartDelayDone = true;
                    }
                }
                SpeedTimer++;
                if (SpeedTimer >= 20)
                {
                    if (!GetDiff)
                    {
                        OldStroke = Stroke;
                        OldForce = Force;
                        OldTime = Time;
                        GetDiff = true;
                    }
                    else
                    {
                        if (Time != OldTime)
                        {
                            Speed = (Stroke - OldStroke) / (Time - OldTime) * 1000;
                            FSpeed = (Force - OldForce);
                        }
                        else
                        {
                            Speed = 0;
                            FSpeed = 0;
                        }
                        GetDiff = false;
                    }
                    SpeedTimer = 0;
                }
                if (((StatusWord & 2) == 2) || ((StatusWord & 4) == 4))
                if (State > 0 && StartDelayDone)
                {
                    if (Speed == 0)
                    {
                        Stop();
                        MessageBox.Show("Encoder Not Detected");
                        return;
                    }
                        if ((Math.Abs(FSpeed) - (SlopeRate / 20)) > 5 || FSpeed == 0)
                    {
                        Stop();
                        MessageBox.Show("Load Cell Not Detected");
                        return;
                    }
                }
                if (Current != 0)
                    CurrentTextBox.Text = Current.ToString("0.0");
                else
                    CurrentTextBox.Text = "0.0";

                CycleTextBox.Text = CycleNumber.ToString();
                ForceTextBox.Text = Force.ToString("0.00");
                StrokeTextBox.Text = Stroke.ToString("0");
                SpeedTextBox.Text = Speed.ToString("0.0");

                if ((StatusWord & 0x4000) == 0x4000)
                {
                    TempError = true;
                    TempStatusLabel.Text = "Temperature Error";
                }
                else
                {
                    TempError = false;
                    TempStatusLabel.Text = "Temperature OK";
                }

                if (((StatusWord & 0x4000) == 0x0) && ((StatusWord & 0x8000) == 0x8000))
                {
                    PowerError = true;
                    PowerStatusLabel.Text = "Main Power Error";
                }
                else
                {
                    PowerError = false;
                    PowerStatusLabel.Text = "Main Power OK";
                }

                StatusLabel.Text = StatusText;
                if (Auto && !StartTest)
                {
                    StopBtn.Hide();
                    StartBtn.Show();
                }
                if (Auto && StartTest)
                {
                    StartBtn.Hide();
                    StopBtn.Show();
                }
                if (StartTest)
                {
                    AutoRadioButton.Enabled = false;
                    ManualRadioButton.Enabled = false;
                    OpenConfigBtn.Enabled = false;
                    ApplyBtn.Enabled = false;
                }
                else
                {
                    AutoRadioButton.Enabled = true;
                    ManualRadioButton.Enabled = true;
                    OpenConfigBtn.Enabled = true;
                }
                if (!Auto)
                {
                    ZeroForceBtn.Show();
                    ZeroPositionBtn.Show();
                    ManualRadioButton.Checked = true;
                    AutoRadioButton.Checked = false;
                    StartBtn.Hide();
                    StopBtn.Hide();
                    VoiceCoilGroupBox.Show();
                }
                else
                {
                    ZeroForceBtn.Hide();
                    ZeroPositionBtn.Hide();
                    AutoRadioButton.Checked = true;
                    ManualRadioButton.Checked = false;
                    VoiceCoilGroupBox.Hide();
                }
                if ((StatusWord & 2) == 2)
                {
                    PositiveLEDOff.Hide();
                    PositiveLEDOn.Show();
                    NegativeBtn.Enabled = false;
                    MaxCurrentTextBox.Text = MaxCurrent.ToString("0.00");
                    MaxForceTextBox.Text = MaxForce.ToString("0.00");
                    MaxStrokeTextBox.Text = MaxStroke.ToString("0");
                }
                else
                {
                    PositiveLEDOn.Hide();
                    PositiveLEDOff.Show();
                    NegativeBtn.Enabled = true;
                }
                if ((StatusWord & 4) == 4)
                {
                    NegativeLEDOff.Hide();
                    NegativeLEDOn.Show();
                    PositiveBtn.Enabled = false;
                    MaxCurrentTextBox.Text = (MaxCurrent * -1).ToString("0.00");
                    MaxForceTextBox.Text = (MaxForce * -1).ToString("0.00");
                    MaxStrokeTextBox.Text = (MaxStroke * -1).ToString("0");
                }
                else
                {
                    NegativeLEDOn.Hide();
                    NegativeLEDOff.Show();
                    PositiveBtn.Enabled = true;
                }
                if ((StatusWord & 6) == 0 && Auto == false)
                {
                    CurrenthScrollBar.Enabled = false;
                }
                else
                    CurrenthScrollBar.Enabled = true;
            }
            else
            {
                if (!Auto)
                {
                    ZeroForceBtn.Show();
                    ZeroPositionBtn.Show();
                    ManualRadioButton.Checked = true;
                    AutoRadioButton.Checked = false;
                    StartBtn.Hide();
                    StopBtn.Hide();
                    VoiceCoilGroupBox.Show();
                }
                else
                {
                    ZeroForceBtn.Hide();
                    ZeroPositionBtn.Hide();
                    AutoRadioButton.Checked = true;
                    ManualRadioButton.Checked = false;
                    VoiceCoilGroupBox.Hide();
                }
                TempStatusLabel.Text = "Temperature State Unknown";
                PowerStatusLabel.Text = "Power State Unknown";
            }

        }
        void SaveConfigFile()
        {
            String text = "";
            lines[7] = "LoadcellOffset=" + LoadOffset.ToString();
            lines[8] = "LoadcellGain=" + LoadGain.ToString() + "\r\n";
            for (int i = 0; i < 9; i++)
            {
                text += lines[i] + "\r\n";
            }
            System.IO.File.WriteAllText(ConfigFilePath, text);
        }
        void ReadConfigFile()
        {
            int i;
            char[] ch = new char[100];
            string[] Lines = new string[10];
            string[] Lines2 = new string[10];
            String[] slopeRateText = new String[10];
            String[] cycleText = new String[10];
            String[] maxCuurentText = new String[10];
            String[] maxForceText = new String[10];
            String[] maxStrokeText = new String[10];
            String[] ScaleFactorText = new String[10];
            String[] LoadOffsetText = new String[10];
            String[] LoadGainText = new String[10];
            string[] PositionOffsetText = new string[10];
            String[] words = new String[10];
            String text = " ";
            char[] charsToTrim = { '\r'};
            try
            {
                text = System.IO.File.ReadAllText(ConfigFilePath);
                text = Regex.Replace(text, "\r", string.Empty);
                lines = text.Split('\n');
                try
                { 
                    for (i = 0; i < lines.Length; i++)              // JJ
                    {
                        if (lines[i].StartsWith("////")) continue;  //JJ
                        words = lines[i].Split('=');
                        switch (words[0])
                        {
                            case "SlopeRate(mA/s)":
                                slopeRateText = words[1].Split(' ');
                                break;
                            case "Cycle#":
                                cycleText = words[1].Split(' ');
                                break;
                            case "MaxCurrent(mA)":
                                maxCuurentText = words[1].Split(' ');
                                break;
                            case "MaxForce(N)":
                                maxForceText = words[1].Split(' ');
                                break;
                            case "MaxStroke(um)":
                                maxStrokeText = words[1].Split(' ');
                                break;
                            case "ScaleFactor":
                                ScaleFactorText = words[1].Split(' ');
                                break;
                            case "PositionOffset":
                                PositionOffsetText = words[1].Split(' ');
                                break;
                            case "LoadcellOffset":
                                LoadOffsetText = words[1].Split(' ');
                                break;
                            case "LoadcellGain":
                                LoadGainText = words[1].Split(' ');
                                break;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Wrong Format");
                }
            }
            catch
            {
                MessageBox.Show("Config File Not Found");
                return;
            }
            try
            {
                if (slopeRateText[0] == null)
                {
                    MessageBox.Show("Invalid Slope Rate");
                }
                else 
                   SlopeRate = Convert.ToDouble(slopeRateText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Slope Rate");
            }
            try
            {
                if (cycleText[0] == null)
                {
                    MessageBox.Show("Invalid Cycle Number");
                }
                else
                    CycleSP = Convert.ToInt32(cycleText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Cycle Number");
            }
            try
            {
                if (maxCuurentText[0] == null)
                {
                    MessageBox.Show("Invalid Max Current");
                }
                else
                {
                    if (Convert.ToDouble(maxCuurentText[0]) <= 6000)
                        MaxCurrent = Convert.ToDouble(maxCuurentText[0]);
                    else
                    {
                        MaxCurrent = 6000;
                    }
                }
            }
            catch
            {
                MessageBox.Show("Invalid Max Current");
            }
            try
            {
                if (maxForceText[0] == null)
                {
                    MessageBox.Show("Invalid Max Force");
                }
                else
                    MaxForce = Convert.ToDouble(maxForceText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Max Force");
            }
            try
            {
                if (maxStrokeText[0] == null)
                {
                    MessageBox.Show("Invalid Max Stroke");
                }
                else
                    MaxStroke = Convert.ToDouble(maxStrokeText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Max Stroke");
            }
            try
            {
                if (ScaleFactorText[0] == null)
                {
                    MessageBox.Show("Invalid Scale Factor");
                }
                else
                    ScaleFactor = Convert.ToDouble(ScaleFactorText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Scale Factor");
            }
            try
            {
                if (PositionOffsetText[0] == null)
                {
                    MessageBox.Show("Invalid Position Offset");
                }
                else
                    PositionOffset = Convert.ToDouble(PositionOffsetText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Position Offset");
            }
            try
            {
                if (LoadOffsetText[0] == null)
                {
                    MessageBox.Show("Invalid Loadcell Offset");
                }
                else
                    LoadOffset = Convert.ToUInt32(LoadOffsetText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Loadcell Offset");
            }
            try
            {
                if (LoadGainText[0] == null)
                {
                    MessageBox.Show("Invalid Loadcell Gain");
                }
                else
                    LoadGain = Convert.ToDouble(LoadGainText[0]);
            }
            catch
            {
                MessageBox.Show("Invalid Loadcell Gain");
            }
        }
        /*
        private void SaveData(int cycleNumber)
        {
            long i;
            long dataNumber1 = 0, dataNumber2 = 0, dataNumber3 = 0, dataNumber4 = 0;
            long n;
            double peakCurrent = 0;
            string delimiter = ",";

            string fileName;
            string[] table = new string[8640000];
            StringBuilder sb = new StringBuilder();

            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = ", , , ,";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            { 
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = 8; i < DataCount; i++)
            {
                table[i] = TimeData[i - 8].ToString("0.00") + "," + ForceData[i - 8].ToString("0.00") + "," + StrokeData[i - 8].ToString("0.00") + "," + CurrentData[i - 8].ToString("0.00") + "," + IndexData[i - 8].ToString();
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test"+ TestNumberTextBox.Text +"_ALL_"+ cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@"c:\Test Report\" + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            n = 0;
            while (n<DataCount)
            {
                if (CurrentData[n] >= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber1 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] <= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber2 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] >= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber3 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] <= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber4 = n;
                    break;
                }
                n++;
            }
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "A";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber1; i < dataNumber2; i++)
            {
                table[i - dataNumber1 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString(); ;
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber1 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_A_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@"c:\Test Report\" + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "B";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber2; i < dataNumber3; i++)
            {
                table[i- dataNumber2 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString(); 
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber2 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_B_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@"c:\Test Report\" + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();

            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "C";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber3; i < dataNumber4; i++)
            {
                table[i- dataNumber3 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString(); 
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber3 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_C_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@"c:\Test Report\" + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "D";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            i = dataNumber4;
            while (i < DataCount && CurrentData[i] != 0)
            {
                table[i - dataNumber4 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString(); 
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber4 + 9]));
                i++;
            }
            table[i - dataNumber4] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00");
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_D_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@"c:\Test Report\" + fileName, sb.ToString());
            
        }
        */

        private void SaveData(int cycleNumber)
        {
            long i;
            long dataNumber1 = 0, dataNumber2 = 0, dataNumber3 = 0, dataNumber4 = 0;
            long n;
            double peakCurrent = 0;
            string delimiter = ",";

            string fileName;
            string[] table = new string[8640000];
            StringBuilder sb = new StringBuilder();

            if (0 == cycleNumber) return; // JJ

            String dirPath = reportPath + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + "\\";
            if (!Directory.Exists(dirPath))
                Directory.CreateDirectory(dirPath);

            try
            {
                Directory.SetCurrentDirectory(dirPath);
            } catch (DirectoryNotFoundException e)
            {
                Console.WriteLine("Directory not found {0}", e);
            }
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = ", , , ,";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = 8; i < DataCount; i++)
            {
                table[i] = TimeData[i - 8].ToString("0.00") + "," + ForceData[i - 8].ToString("0.00") + "," + StrokeData[i - 8].ToString("0.00") + "," + CurrentData[i - 8].ToString("0.00") + "," + IndexData[i - 8].ToString();
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_ALL_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@dirPath + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            n = 0;
            while (n < DataCount)
            {
                if (CurrentData[n] >= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber1 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] <= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber2 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] >= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber3 = n;
                    break;
                }
                n++;
            }
            while (n < DataCount)
            {
                if (CurrentData[n] <= peakCurrent)
                {
                    peakCurrent = CurrentData[n];
                }
                else
                {
                    dataNumber4 = n;
                    break;
                }
                n++;
            }
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "A";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber1; i < dataNumber2; i++)
            {
                table[i - dataNumber1 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString(); ;
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber1 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_A_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@dirPath + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "B";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber2; i < dataNumber3; i++)
            {
                table[i - dataNumber2 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString();
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber2 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_B_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@dirPath + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();

            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "C";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            for (i = dataNumber3; i < dataNumber4; i++)
            {
                table[i - dataNumber3 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString();
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber3 + 9]));
            }
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_C_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@dirPath + fileName, sb.ToString());
            Array.Clear(table, 0, table.Length);
            sb.Clear();
            table[0] = "Operator Name," + OperatorTextBox.Text + ", ,";
            table[1] = "Organization name," + OrganizationTextBox.Text + ", ,";
            table[2] = "Test Date," + Date;
            table[3] = "Test Number," + TestNumberTextBox.Text + ", ,";
            table[4] = "Serial Number," + SerialNumberTextBox.Text + ", ,";
            table[5] = "Notes," + NotesTextBox.Text + ", ,";
            table[6] = "D";
            table[7] = "Time(mS),Force(N),Stroke(Micron),Current(mA),Index";
            for (i = 0; i < 8; i++)
            {
                sb.AppendLine(string.Join(delimiter, table[i]));
            }
            i = dataNumber4;
            while (i < DataCount && CurrentData[i] != 0)
            {
                table[i - dataNumber4 + 9] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00") + "," + IndexData[i].ToString();
                sb.AppendLine(string.Join(delimiter, table[i - dataNumber4 + 9]));
                i++;
            }
            table[i - dataNumber4] = TimeData[i].ToString("0.00") + "," + ForceData[i].ToString("0.00") + "," + StrokeData[i].ToString("0.00") + "," + CurrentData[i].ToString("0.00");
            fileName = "SN" + SerialNumberTextBox.Text + "_Test" + TestNumberTextBox.Text + "_D_" + cycleNumber.ToString() + "_Date" + TimeStamp + ".CSV";
            System.IO.File.WriteAllText(@dirPath + fileName, sb.ToString());

            // JJ ZipFile.CreateFromDirectory(dirPath, basePath + DateTime.Now("yyyyy-MM-dd-hh-mm"));
        }
        private void AutoMode(bool mode)
        {
            if (mode)
            {
                ControlWord = 1;
                SendControl = true;
            }
            else
            {
                ControlWord = 0;
                SendControl = true;
            }
        }
        private void PositiveCurrentOn(bool on)
        {
            if (on)
            {
                ControlWord &= 0xF3;
                ControlWord |= 2;
                SendControl = true;
                CurrenthScrollBar.Value = 0;
            }
            else
            {
                ControlWord &= 0xFD;
                SendControl = true;
            }
        }
        private void NegativeCurrentOn(bool on)
        {
            if (on)
            {
                ControlWord &= 0xF5;
                ControlWord |= 4;
                SendControl = true;
                CurrenthScrollBar.Value = 0;
            }
            else
            {
                ControlWord &= 0xFB;
                SendControl = true;
            }
        }
        private void CurrentOff()
        {
            ControlWord &= 0xCF;
            ControlWord |= 0x8;
            SendControl = true;
            CurrenthScrollBar.Value = 0;
        }

        private void Stop()
        {
            CurrentOff();
            CollectData = false;
            StartTest = false;
            SaveData(CycleNumber);
            StartTime = 0;
            StartDelayDone = false;
            SpeedTimer = 0;
        }
        private void Start()
        {
            ControlWord = 0x91;
            SendControl = true;
            DataReceived = false;
            ForceStrokeChart.Series[0].Points.Clear();
            CurrentStrokeChart.Series[0].Points.Clear();
            ForceTimeChart.Series[0].Points.Clear();
            DataCount = 0;
            TimeStamp = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
            Date = DateTime.Now.ToString("yyyy-MM-dd");
            CollectData = true;
            StartTest = true;
            CycleNumber = 1;
            StartTime = GetTime();
            ProgressBar.Value = 0;
            OldRawTime = 0;
            SpeedTimer = 0;
            StartDelay = 0;
            StartDelayDone = false;
            

            this.ForceStrokeChart.Series[0].MarkerSize = 6;
            this.ForceStrokeChart.Series[0].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Cross;
            this.ForceStrokeChart.Series[0].MarkerColor = Color.Black;

            this.CurrentStrokeChart.Series[0].MarkerSize = 6;
            this.CurrentStrokeChart.Series[0].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Cross;
            this.CurrentStrokeChart.Series[0].MarkerColor = Color.Black;

            this.ForceTimeChart.Series[0].MarkerSize = 6;
            this.ForceTimeChart.Series[0].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Cross;
            this.ForceTimeChart.Series[0].MarkerColor = Color.Black;

            this.ForceStrokeChart.Series[0].Color = Color.Magenta;
            this.CurrentStrokeChart.Series[0].Color = Color.Blue;
            this.ForceTimeChart.Series[0].Color = Color.Green;

            for (int i = 0;i< 8640000; i++)
            {
                IndexData[i] = 0;
            }

            ForceStrokeChart.Show();
            CurrentStrokeChart.Show();
            ForceTimeChart.Show();
        }
        private void Restart()
        {
            ControlWord = 0x91;
            SendControl = true;
            DataReceived = false;
            ForceStrokeChart.Series[0].Points.Clear();
            CurrentStrokeChart.Series[0].Points.Clear();
            ForceTimeChart.Series[0].Points.Clear();
            DataCount = 0;
            CollectData = true;
            StartTest = true;
            StartTime = GetTime();
            OldRawTime = 0;
        }
        private void ZeroPosition()
        {
            ControlWord |= 64;
            SendControl = true;
        }
        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (!OfflineMode) Start();
        }
        private void StopBtn_Click(object sender, EventArgs e)
        {
            if (!OfflineMode) Stop();
        }
        private void LoadData()
        {
            MaxRawCurrent = (long)(((MaxCurrent/CurrentGain) * 262144) / 20000);
            RampFreq = (short)(SlopeRate * SpeedRatio);
            CurrenthScrollBar.Maximum = (int)MaxRawCurrent;
            if (RampFreq > 0)
            {
                PR2 = (273437 / RampFreq);   //Prescale = 256
                RampTimeMultiply = 1000 / (double)RampFreq;
                TimeRemaining = MaxRawCurrent * RampTimeMultiply * 0.001 * 8 * CycleSP;
                ProgressBar.Maximum = (int)TimeRemaining;
            }
            SendCofig = true;
            ApplyBtn.Enabled = false;
            SlopeRateTextBox.Text = SlopeRate.ToString("0.0");
            CycleSPTextBox.Text = CycleSP.ToString();
            MaxCurrentTextBox.Text = MaxCurrent.ToString("0.00");
            MaxForceTextBox.Text = MaxForce.ToString("0.00");
            MaxStrokeTextBox.Text = MaxStroke.ToString("0");
            TimeSpan t = TimeSpan.FromSeconds(TimeRemaining);
            TimeRemainingTextBox.Text = t.Minutes.ToString() + ":" + t.Seconds.ToString();
        }
        private void ApplyBtn_Click(object sender, EventArgs e)
        {
            ReadConfigFile();
            LoadData();
        }
        private void ManualRadioButton_MouseDown(object sender, MouseEventArgs e)
        {
            AutoMode(false);
            Auto = false;
        }
        private void ZeroForceBtn_Click(object sender, EventArgs e)
        {
            if (!OfflineMode)
            {
               LoadOffset = RawForce;
               SaveConfigFile();
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ReadConfigFile();
            ConfigFilePathLabel.Text = ConfigFilePath;
            AutoGroupBox.Hide();
            AutoMode(true);
            Auto = true;
            Stop();
            if (OfflineMode) LoadDataOnce = true;
        }
        private void DownOnBtn_Click_1(object sender, EventArgs e)
        {
            NegativeCurrentOn(true);
        }



        private void ZeroPositionBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Confirm Fixture Bar is in Place", "Important Question", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (!OfflineMode) ZeroPosition();
            }
            
        }
        private void OpenConfigBtn_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(ConfigFilePath))
            {
                System.Diagnostics.Process.Start("notepad.exe", ConfigFilePath);
                ApplyBtn.Enabled = true;
            }
            else
            {
                MessageBox.Show("Config File Not Found");
                ApplyBtn.Enabled = false;
            }
        }
        private void AutoRadioButton_MouseDown(object sender, MouseEventArgs e)
        {
            AutoMode(true);
            Auto = true;
        }
        private void ForceStrokeRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (ForceStrokeRadioButton.Checked) 
            {
                ForceStrokeChart.Show();
                CurrentStrokeChart.Hide();
                ForceTimeChart.Hide();
            }
        }

        private void ForceTimeRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (ForceTimeRadioButton.Checked)
            {
                ForceTimeChart.Show();
                ForceStrokeChart.Hide();
                CurrentStrokeChart.Hide();
            }
        }

        private void OpenLogBtn_Click(object sender, EventArgs e)
        {
            SaveData(CycleNumber);
        }

        private void CurrentStrokeRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (CurrentStrokeRadioButton.Checked)
            {
                CurrentStrokeChart.Show();
                ForceStrokeChart.Hide();
                ForceTimeChart.Hide();
            }

        }

        private void OperatorRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (OperatorRadioButton.Checked)
            {
                AutoGroupBox.Hide();
                StartBtn.Show();
                StopBtn.Hide();
                AutoMode(true);
                Auto = true;
            }
        }

        private void EngineerRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (EngineerRadioButton.Checked)
            {
                PasswordGroupBox.Show();
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            PasswordGroupBox.Hide();
            OperatorRadioButton.Checked = true;
            EngineerRadioButton.Checked = false;

        }
        private void CheckPassword()
        {
            DialogResult result;
            if (PasswordTextBox.Text == Password)
            {
                AutoGroupBox.Show();
                PasswordGroupBox.Hide();
            }
            else
            {
                OperatorRadioButton.Checked = true;
                EngineerRadioButton.Checked = false;
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                result = MessageBox.Show("Incorrect Password", "", buttons);
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    PasswordTextBox.Clear();
                }

            }
        }
        private void OKButton_Click(object sender, EventArgs e)
        {
            CheckPassword();
        }

        private void CurrenthScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            if (!OfflineMode)
            {
                AnalogSP = CurrenthScrollBar.Value;
                SendCofig = true;
            }
        }
        private void PositiveBtn_Click(object sender, EventArgs e)
        {
            if (!OfflineMode) PositiveCurrentOn(true);
        }

        private void NegativeBtn_Click(object sender, EventArgs e)
        {
            if (!OfflineMode) NegativeCurrentOn(true);
        }


        private void OffBtn_Click(object sender, EventArgs e)
        {
            CurrentOff();
        }

        private void PasswordTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                CheckPassword();
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.TMTWebURL);
        }
    }
}
