using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace BluetoothBatteryTest
{

    class SetupAPI
    {
        //https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetclassdevsw
        [DllImport(@"Setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetupDiGetClassDevsW(
            ref Guid ClassGuid, 
            String Enumerator, 
            IntPtr hwndParent, 
            UInt32 Flags
        );


        //https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdienumdeviceinterfaces
        [DllImport(@"Setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern Boolean SetupDiEnumDeviceInterfaces(
            IntPtr DeviceInfoSet, 
            IntPtr DeviceInfoData, 
            ref Guid InterfaceClassGuid, 
            Int32 MemberIndex, 
            ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData
        );


        //https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetdeviceinterfacedetailw
        [DllImport(@"Setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern Boolean SetupDiGetDeviceInterfaceDetailW(
            IntPtr DeviceInfoSet, 
            ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData, 
            ref SP_DEVICE_INTERFACE_DETAIL_DATA DeviceInterfaceDetailData, 
            Int32 DeviceInterfaceDetailDataSize, 
            ref UInt32 RequiredSize, 
            ref SP_DEVINFO_DATA DeviceInfoData
        );


        //https://docs.microsoft.com/en-us/windows/win32/api/setupapi/ns-setupapi-sp_devinfo_data
        //https://www.pinvoke.net/default.aspx/Structures/SP_DEVINFO_DATA.html
        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DEVINFO_DATA
        {
            public Int32 cbSize;
            public Guid ClassGuid;
            public UInt32 DevInst;
            public IntPtr Reserved;
        }


        //https://docs.microsoft.com/en-us/windows/win32/api/setupapi/ns-setupapi-sp_device_interface_data
        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DEVICE_INTERFACE_DATA
        {
            public Int32 cbSize;
            public Guid interfaceClassGuid;
            public UInt32 flags;
            private UIntPtr reserved;
        }


        //https://docs.microsoft.com/zh-cn/windows/win32/api/setupapi/ns-setupapi-sp_device_interface_detail_data_w
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SP_DEVICE_INTERFACE_DETAIL_DATA
        {
            public Int32 cbSize;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string DevicePath;
        }


        public static readonly UInt32 DIGCF_PRESENT = 0x02;
        public static readonly UInt32 DIGCF_DEVINTERFACE = 0x10;
    }

    class Kernel32
    {
        //https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew
        [DllImport("Kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CreateFileW(
            [MarshalAs(UnmanagedType.LPWStr)] string filename,
            [MarshalAs(UnmanagedType.U4)] FileAccess access,
            [MarshalAs(UnmanagedType.U4)] FileShare share,
            IntPtr securityAttributes,
            [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
            [MarshalAs(UnmanagedType.U4)] FileAttributes flagsAndAttributes,
            IntPtr templateFile
        );
    }

    class BluetoothAPIs
    {
        //https://docs.microsoft.com/zh-cn/windows/win32/api/bluetoothleapis/nf-bluetoothleapis-bluetoothgattgetservices
        [DllImport("BluetoothAPIs.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int BluetoothGATTGetServices(
            IntPtr hDevice,
            UInt16 ServicesBufferCount,
            IntPtr ServicesBuffer,
            ref UInt16 ServicesBufferActual,
            UInt32 Flags
        );

        //https://docs.microsoft.com/zh-cn/windows/win32/api/bthledef/ns-bthledef-bth_le_uuid
        [StructLayout(LayoutKind.Explicit)]
        public struct BTH_LE_UUID
        {
            [FieldOffset(0)]
            public Byte IsShortUuid;
            [FieldOffset(4)]
            public UInt16 ShortUuid;
            [FieldOffset(4)]
            public Guid LongUuid;
        }

        //https://docs.microsoft.com/zh-cn/windows/win32/api/bthledef/ns-bthledef-bth_le_gatt_service
        [StructLayout(LayoutKind.Sequential)]
        public struct BTH_LE_GATT_SERVICE
        {
            public BTH_LE_UUID ServiceUuid;
            public UInt16 AttributeHandle;
        }

        //https://docs.microsoft.com/en-us/windows/win32/api/bluetoothleapis/nf-bluetoothleapis-bluetoothgattgetcharacteristics
        [DllImport("BluetoothAPIs.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int BluetoothGATTGetCharacteristics(
            IntPtr hDevice,
            ref BTH_LE_GATT_SERVICE Service,
            UInt16 CharacteristicsBufferCount,
            IntPtr CharacteristicsBuffer,
            ref UInt16 CharacteristicsBufferActual,
            UInt32 Flags
        );

        //https://docs.microsoft.com/en-us/windows/win32/api/bthledef/ns-bthledef-bth_le_gatt_characteristic
        [StructLayout(LayoutKind.Sequential)]
        public struct BTH_LE_GATT_CHARACTERISTIC
        {
            public UInt16 ServiceHandle;
            public BTH_LE_UUID CharacteristicUuid;
            public UInt16 AttributeHandle;
            public UInt16 CharacteristicValueHandle;
            public Byte IsBroadcastable;
            public Byte IsReadable;
            public Byte IsWritable;
            public Byte IsWritableWithoutResponse;
            public Byte IsSignedWritable;
            public Byte IsNotifiable;
            public Byte IsIndicatable;
            public Byte HasExtendedProperties;
        }

        //https://docs.microsoft.com/en-us/windows/win32/api/bluetoothleapis/nf-bluetoothleapis-bluetoothgattgetcharacteristicvalue
        [DllImport("BluetoothAPIs.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int BluetoothGATTGetCharacteristicValue(
            IntPtr hDevice,
            ref BTH_LE_GATT_CHARACTERISTIC Characteristic,
            UInt32 CharacteristicValueDataSize,
            ref BTH_LE_GATT_CHARACTERISTIC_VALUE CharacteristicValue,
            out UInt16 CharacteristicValueSizeRequired,
            UInt32 Flags
        );

        //https://docs.microsoft.com/en-us/windows/win32/api/bthledef/ns-bthledef-bth_le_gatt_characteristic_value
        [StructLayout(LayoutKind.Sequential)]
        public struct BTH_LE_GATT_CHARACTERISTIC_VALUE
        {
            public UInt32 DataSize;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)] public byte[] Data;
        }

        public static readonly UInt32 BLUETOOTH_GATT_FLAG_NONE = 0;
        public static readonly UInt32 BLUETOOTH_GATT_FLAG_FORCE_READ_FROM_DEVICE = 0x00000004;
    }

    class Program
    {
        //GATT Service UUIDs 
        private static readonly string UUID1 = "{0000180F-0000-1000-8000-00805F9B34FB}"; //Battery Service
        private static readonly string UUID2 = "{00001800-0000-1000-8000-00805F9B34FB}"; //appearance
        private static readonly string UUID3 = "{0000180A-0000-1000-8000-00805F9B34FB}"; //DEVICE INFORMATION

        

        private static readonly Dictionary<int, string> appearanceDict = new Dictionary<int, string>(){
            { 64, "Phone"}, {961, "Keyboard"}, {962, "Mouse"}, {963, "Joystick"}, {964, "Gamepad"} 
        };

        static void Main(string[] args)
        {
            getDevInfo(Guid.Parse(UUID1));
            getDevInfo(Guid.Parse(UUID2));
            getDevInfo(Guid.Parse(UUID3));
        }

        static void getDevInfo(Guid guid)
        {
            IntPtr hdi = SetupAPI.SetupDiGetClassDevsW(ref guid, null, IntPtr.Zero, SetupAPI.DIGCF_PRESENT | SetupAPI.DIGCF_DEVINTERFACE);
            Console.WriteLine(hdi);
            SetupAPI.SP_DEVINFO_DATA dd = new SetupAPI.SP_DEVINFO_DATA();
            dd.cbSize = Marshal.SizeOf(dd);
            SetupAPI.SP_DEVICE_INTERFACE_DATA did = new SetupAPI.SP_DEVICE_INTERFACE_DATA();
            did.cbSize = Marshal.SizeOf(did);
            int i = 0;
            while (SetupAPI.SetupDiEnumDeviceInterfaces(hdi, IntPtr.Zero, ref guid, i, ref did))
            {
                i += 1;
                SetupAPI.SP_DEVICE_INTERFACE_DETAIL_DATA didd = new SetupAPI.SP_DEVICE_INTERFACE_DETAIL_DATA();
                didd.cbSize = IntPtr.Size == 8 ? 8 : 4 + Marshal.SystemDefaultCharSize;
                UInt32 nRequiredSize = 0;
                SetupAPI.SetupDiGetDeviceInterfaceDetailW(hdi, ref did, ref didd, 256, ref nRequiredSize, ref dd);
                Console.WriteLine(didd.DevicePath);
                communication(didd.DevicePath);
            }
            Console.WriteLine("\n\n");
        }

        static void communication(string path)
        {
            IntPtr hDevice = Kernel32.CreateFileW(path, FileAccess.Read, FileShare.ReadWrite, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            UInt16 servicesBufferCount = 0;
            BluetoothAPIs.BluetoothGATTGetServices(hDevice, 0, IntPtr.Zero, ref servicesBufferCount, BluetoothAPIs.BLUETOOTH_GATT_FLAG_NONE);
            Console.WriteLine("Fount {0} services", servicesBufferCount);
            IntPtr servicesBufferPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(BluetoothAPIs.BTH_LE_GATT_SERVICE)) * servicesBufferCount);
            BluetoothAPIs.BluetoothGATTGetServices(hDevice, servicesBufferCount, servicesBufferPtr, ref servicesBufferCount, BluetoothAPIs.BLUETOOTH_GATT_FLAG_NONE);
            for (int i = 0; i < servicesBufferCount; i++)
            {
                BluetoothAPIs.BTH_LE_GATT_SERVICE service = Marshal.PtrToStructure<BluetoothAPIs.BTH_LE_GATT_SERVICE>(servicesBufferPtr + i * Marshal.SizeOf(typeof(BluetoothAPIs.BTH_LE_GATT_SERVICE)));
                UInt16 characteristicsBufferCount = 0;
                int hr = BluetoothAPIs.BluetoothGATTGetCharacteristics(hDevice, ref service, 0, IntPtr.Zero, ref characteristicsBufferCount, BluetoothAPIs.BLUETOOTH_GATT_FLAG_NONE);
                Console.WriteLine("ResultCode: {0}, Fount {1} characteristics", hr, characteristicsBufferCount);
                IntPtr characteristicsBufferPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC)) * characteristicsBufferCount);
                UInt16 characteristicsBufferActual = 0;
                BluetoothAPIs.BluetoothGATTGetCharacteristics(hDevice, ref service, characteristicsBufferCount, characteristicsBufferPtr, ref characteristicsBufferActual, BluetoothAPIs.BLUETOOTH_GATT_FLAG_NONE);
                for (int n = 0; n < characteristicsBufferActual; n++)
                {
                    BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC characteristic = Marshal.PtrToStructure<BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC>(characteristicsBufferPtr + n * Marshal.SizeOf(typeof(BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC)));
                    BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC_VALUE characteristicValue = new BluetoothAPIs.BTH_LE_GATT_CHARACTERISTIC_VALUE();
                    UInt16 characteristicValueSizeRequired = 0;
                    BluetoothAPIs.BluetoothGATTGetCharacteristicValue(hDevice, ref characteristic, 256, ref characteristicValue, out characteristicValueSizeRequired, BluetoothAPIs.BLUETOOTH_GATT_FLAG_NONE);
                    if (characteristic.CharacteristicUuid.ShortUuid == 0x2A19) // battery level
                    {
                        Console.WriteLine("Battery Level: {0}", characteristicValue.Data[0]);
                    }
                    else if (characteristic.CharacteristicUuid.ShortUuid == 0x2A01) // appearance
                    {
                        int appearance = BitConverter.ToInt32(characteristicValue.Data);
                        string appearanceName = appearanceDict.GetValueOrDefault(appearance, "Unknown");
                        Console.WriteLine("Appearance: {0}, AppearanceName: {1}", appearance, appearanceName);
                    }
                    else //string
                    {
                        Console.WriteLine("IsShortUuid：{0}， UUID: {1}, DataSize: {2}, Data: {3}", characteristic.CharacteristicUuid.IsShortUuid, characteristic.CharacteristicUuid.ShortUuid, characteristicValue.DataSize, Encoding.ASCII.GetString(characteristicValue.Data, 0, (int)characteristicValue.DataSize));
                    }
                }
                Marshal.FreeHGlobal(characteristicsBufferPtr);
            }
            Marshal.FreeHGlobal(servicesBufferPtr);
        }
    }
}
