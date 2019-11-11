using System;
using System.Runtime.InteropServices;
using System.Management;

namespace OfficeAssist
{
    /// <summary> 
    /// Hardware_Mac 的摘要说明。 
    /// </summary> 
    public class ClassHardInfo
    {
        //取CPU编号 
        public string GetCpuID()
        {
            try
            {
                ManagementClass mc = new ManagementClass("Win32_Processor");
                ManagementObjectCollection moc = mc.GetInstances();

                string strCpuID = null;
                foreach (ManagementObject mo in moc)
                {
                    strCpuID = mo.Properties["ProcessorId"].Value.ToString();
                    break;
                }
                return strCpuID;
            }
            catch
            {
                return "";
            }

        }//end method 

        public string GetDiskID()
        {
            try
            {
                //获取硬盘ID 
                String HDid = "";
                ManagementClass mc = new ManagementClass("Win32_DiskDrive");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    HDid = (string)mo.Properties["Model"].Value;
                }
                moc = null;
                mc = null;
                return HDid;
            }
            catch
            {
                return "";
            }
            finally
            {
            }

        }


        //取第一块硬盘编号 
        public string GetHardDiskID()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");
                string strHardDiskID = null;
                foreach (ManagementObject mo in searcher.Get())
                {
                    strHardDiskID = mo["SerialNumber"].ToString().Trim();
                    break;
                }
                return strHardDiskID;
            }
            catch (Exception ex)
            {
                return "";
            }
        }//end 

        public enum NCBCONST
        {
            NCBNAMSZ = 16, /* absolute length of a net name */
            MAX_LANA = 254, /* lana's in range 0 to MAX_LANA inclusive */
            NCBENUM = 0x37, /* NCB ENUMERATE LANA NUMBERS */
            NRC_GOODRET = 0x00, /* good return */
            NCBRESET = 0x32, /* NCB RESET */
            NCBASTAT = 0x33, /* NCB ADAPTER STATUS */
            NUM_NAMEBUF = 30, /* Number of NAME's BUFFER */
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ADAPTER_STATUS
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
            internal byte[] adapter_address;
            internal byte rev_major;
            internal byte reserved0;
            internal byte adapter_type;
            internal byte rev_minor;
            internal ushort duration;
            internal ushort frmr_recv;
            internal ushort frmr_xmit;
            internal ushort iframe_recv_err;
            internal ushort xmit_aborts;
            internal uint xmit_success;
            internal uint recv_success;
            internal ushort iframe_xmit_err;
            internal ushort recv_buff_unavail;
            internal ushort t1_timeouts;
            internal ushort ti_timeouts;
            internal uint reserved1;
            internal ushort free_ncbs;
            internal ushort max_cfg_ncbs;
            internal ushort max_ncbs;
            internal ushort xmit_buf_unavail;
            internal ushort max_dgram_size;
            internal ushort pending_sess;
            internal ushort max_cfg_sess;
            internal ushort max_sess;
            internal ushort max_sess_pkt_size;
            internal ushort name_count;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct NAME_BUFFER
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = (int)NCBCONST.NCBNAMSZ)]
            internal byte[] name;
            internal byte name_num;
            internal byte name_flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct NCB
        {
            internal byte ncb_command;
            internal byte ncb_retcode;
            internal byte ncb_lsn;
            internal byte ncb_num;
            internal IntPtr ncb_buffer;
            internal ushort ncb_length;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = (int)NCBCONST.NCBNAMSZ)]
            internal byte[] ncb_callname;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = (int)NCBCONST.NCBNAMSZ)]
            internal byte[] ncb_name;
            internal byte ncb_rto;
            internal byte ncb_sto;
            internal IntPtr ncb_post;
            internal byte ncb_lana_num;
            internal byte ncb_cmd_cplt;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
            internal byte[] ncb_reserve;
            internal IntPtr ncb_event;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LANA_ENUM
        {
            internal byte length;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = (int)NCBCONST.MAX_LANA)]
            internal byte[] lana;
        }

        [StructLayout(LayoutKind.Auto)]
        public struct ASTAT
        {
            internal ADAPTER_STATUS adapt;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = (int)NCBCONST.NUM_NAMEBUF)]
            internal NAME_BUFFER[] NameBuff;
        }

        public class Win32API
        {
            [DllImport("NETAPI32.DLL")]
            internal static extern char Netbios(ref NCB ncb);
        }

        public string GetMacAddress()
        {
            string addr = "";
            try
            {
                int cb;
                ASTAT adapter;
                NCB Ncb = new NCB();
                char uRetCode;
                LANA_ENUM lenum;

                Ncb.ncb_command = (byte)NCBCONST.NCBENUM;
                cb = Marshal.SizeOf(typeof(LANA_ENUM));
                Ncb.ncb_buffer = Marshal.AllocHGlobal(cb);
                Ncb.ncb_length = (ushort)cb;
                uRetCode = Win32API.Netbios(ref Ncb);
                lenum = (LANA_ENUM)Marshal.PtrToStructure(Ncb.ncb_buffer, typeof(LANA_ENUM));
                Marshal.FreeHGlobal(Ncb.ncb_buffer);
                if (uRetCode != (short)NCBCONST.NRC_GOODRET)
                    return "";

                for (int i = 0; i < lenum.length; i++)
                {
                    Ncb.ncb_command = (byte)NCBCONST.NCBRESET;
                    Ncb.ncb_lana_num = lenum.lana[i];
                    uRetCode = Win32API.Netbios(ref Ncb);
                    if (uRetCode != (short)NCBCONST.NRC_GOODRET)
                        return "";

                    Ncb.ncb_command = (byte)NCBCONST.NCBASTAT;
                    Ncb.ncb_lana_num = lenum.lana[i];
                    Ncb.ncb_callname[0] = (byte)'*';
                    cb = Marshal.SizeOf(typeof(ADAPTER_STATUS)) + Marshal.SizeOf(typeof(NAME_BUFFER)) * (int)NCBCONST.NUM_NAMEBUF;
                    Ncb.ncb_buffer = Marshal.AllocHGlobal(cb);
                    Ncb.ncb_length = (ushort)cb;
                    uRetCode = Win32API.Netbios(ref Ncb);
                    adapter.adapt = (ADAPTER_STATUS)Marshal.PtrToStructure(Ncb.ncb_buffer, typeof(ADAPTER_STATUS));
                    Marshal.FreeHGlobal(Ncb.ncb_buffer);

                    if (uRetCode == (short)NCBCONST.NRC_GOODRET)
                    {
                        if (i > 0)
                            addr += ":";
                        addr = string.Format("{0,2:X}{1,2:X}{2,2:X}{3,2:X}{4,2:X}{5,2:X}",
                            adapter.adapt.adapter_address[0],
                            adapter.adapt.adapter_address[1],
                            adapter.adapt.adapter_address[2],
                            adapter.adapt.adapter_address[3],
                            adapter.adapt.adapter_address[4],
                            adapter.adapt.adapter_address[5]);
                    }
                }
            }
            catch
            { }
            return addr.Replace(' ', '0');
        }

    }
}
