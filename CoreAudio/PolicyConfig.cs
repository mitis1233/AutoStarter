using System;
using System.Runtime.InteropServices;

namespace AutoStarter.CoreAudio
{
    [ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class _PolicyConfigClient
    {
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    internal interface IPolicyConfig
    {
        [PreserveSig]
        int GetMixFormat(string pszDeviceName, out IntPtr ppFormat);

        [PreserveSig]
        int GetDeviceFormat(string pszDeviceName, bool bDefault, out IntPtr ppFormat);

        [PreserveSig]
        int ResetDeviceFormat(string pszDeviceName);

        [PreserveSig]
        int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr pMixFormat);

        [PreserveSig]
        int GetProcessingPeriod(string pszDeviceName, bool bDefault, out long pmftDefaultPeriod, out long pmftMinPeriod);

        [PreserveSig]
        int SetProcessingPeriod(string pszDeviceName, ref long pmftPeriod);

        [PreserveSig]
        int GetShareMode(string pszDeviceName, out IntPtr pShareMode);

        [PreserveSig]
        int SetShareMode(string pszDeviceName, IntPtr pShareMode);

        [PreserveSig]
        int GetPropertyValue(string pszDeviceName, ref PROPERTYKEY key, out PROPVARIANT pv);

        [PreserveSig]
        int SetPropertyValue(string pszDeviceName, ref PROPERTYKEY key, ref PROPVARIANT propvar);

        [PreserveSig]
        int SetDefaultEndpoint(string pszDeviceName, ERole role);

        [PreserveSig]
        int SetEndpointVisibility(string pszDeviceName, bool bVisible);
    }

    public class PolicyConfigClient
    {
        private readonly IPolicyConfig _policyConfig;

        public PolicyConfigClient()
        {
            _policyConfig = new _PolicyConfigClient() as IPolicyConfig;
        }

        public void SetDefaultEndpoint(string deviceId, ERole role)
        {
            Marshal.ThrowExceptionForHR(_policyConfig.SetDefaultEndpoint(deviceId, role));
        }

        public void DisableEndpoint(string deviceId)
        {
            Marshal.ThrowExceptionForHR(_policyConfig.SetEndpointVisibility(deviceId, false));
        }

        public void EnableEndpoint(string deviceId)
        {
            Marshal.ThrowExceptionForHR(_policyConfig.SetEndpointVisibility(deviceId, true));
        }
    }

    public enum ERole
    {
        eConsole = 0,
        eMultimedia = 1,
        eCommunications = 2,
        ERole_count = 3
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPVARIANT
    {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public IntPtr p;
    }
}
