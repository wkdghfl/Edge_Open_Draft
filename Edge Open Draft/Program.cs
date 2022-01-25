using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;


namespace Edge_Open_Draft
{
    class Program
    {
        public static class Marshal2
        {
            internal const String OLEAUT32 = "oleaut32.dll";
            internal const String OLE32 = "ole32.dll";

            [System.Security.SecurityCritical]  // auto-generated_required
            public static Object GetActiveObject(String progID)
            {
                Object obj = null;
                Guid clsid;

                // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
                // CLSIDFromProgIDEx doesn't exist.
                try
                {
                    CLSIDFromProgIDEx(progID, out clsid);
                }
                //            catch
                catch (Exception)
                {
                    CLSIDFromProgID(progID, out clsid);
                }

                GetActiveObject(ref clsid, IntPtr.Zero, out obj);
                return obj;
            }

            //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
            [DllImport(OLE32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [System.Security.SecurityCritical]  // auto-generated
            private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

            //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
            [DllImport(OLE32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [System.Security.SecurityCritical]  // auto-generated
            private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

            //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
            [DllImport(OLEAUT32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [System.Security.SecurityCritical]  // auto-generated
            private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

        }
        
        // Edge 실행 체크
        static bool IsEdgeStart()
        {
            Process[] processList = Process.GetProcessesByName("Edge");
            if (processList.Length > 0)
                return true;
            return false;
        }

        [STAThread]
        static void Main(string[] args)
        {
            #region 변수선언
            SolidEdgeFramework.Application EdApp = null;
            SolidEdgePart.PartDocument EdPart = null;
            SolidEdgeAssembly.AssemblyDocument EdAssy = null;
            SolidEdgeDraft.DraftDocument EdDft = null;
            #endregion


            try
            {

                //OleMessageFilter.Register();

                #region Edge 연결
                if (IsEdgeStart())
                {
                    EdApp = (Application)Marshal2.GetActiveObject("SolidEdge.Application");

                    String TStr = String.Empty;
                    String TStrDrf = String.Empty;

                    if (EdApp.ActiveSelectSet.Count > 0)
                    {
                        TStr = EdApp.ActiveSelectSet.Item(1).Object.OccurrenceFileName;

                    }
                    else
                    {
                        TStr = EdApp.ActiveDocument.FullName;

                    }

                    TStrDrf = TStr.Substring(0,TStr.Length - 3) + "dft" ;

                    //static string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    //static string directoryPath = Path.GetDirectoryName(exePath);
                    //if (File.Exists(directoryPath + @"\CatiaLubeGroove.pdf")) 
                    if (File.Exists(TStrDrf))
                    {
                        EdApp.Documents.Open(TStrDrf);
                    }
                    else
                    {
                        EdApp.StartCommand("57637");

                    }
                }
                else
                {
                    throw new Exception("SolidEdge 프로그램이 실행되어있지 않습니다.");
                }
                #endregion


            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //OleMessageFilter.Revoke();
            }
            //Console.WriteLine("Hello World!");
        }
    }
}
