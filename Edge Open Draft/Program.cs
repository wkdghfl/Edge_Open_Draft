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
            //SolidEdgePart.PartDocument EdPart = null;
            //SolidEdgeAssembly.AssemblyDocument EdAssy = null;
            //SolidEdgeDraft.DraftDocument EdDft = null;
            #endregion


            try
            {

                //OleMessageFilter.Register();

                #region Edge 연결
                if (IsEdgeStart())
                {
                    EdApp = (Application)Marshal2.GetActiveObject("SolidEdge.Application");
                    String TStr = String.Empty;

                    // 1. 선택된 객체 유무에 따라 원본 파일 경로 가져오기
                    if (EdApp.ActiveSelectSet.Count > 0)
                    {
                        if (EdApp.ActiveSelectSet.Item(1) is SolidEdgeAssembly.Occurrence)
                        {
                            TStr = EdApp.ActiveSelectSet.Item(1).OccurrenceFileName;
                        }
                        else
                        {
                            TStr = EdApp.ActiveSelectSet.Item(1).Object.OccurrenceFileName;
                        }
                    }
                    else
                    {
                        TStr = EdApp.ActiveDocument.FullName;
                    } 

                    // 2. 패밀리 맴버 체크 후 정리
                    if (TStr.Contains("!"))
                    {
                        // "!"를 기준으로 문자를 쪼갠 뒤, 첫 번째 배열([0], 즉 진짜 경로)만 취함
                        TStr = TStr.Split('!')[0];
                    }

                    // C# 기본 클래스를 이용해 안전하게 .dft로 변경 (문자열 길이 계산 오류 방지)
                    TStr = System.IO.Path.ChangeExtension(TStr, ".dft");

                    
                    // 3. 파일 존재 확인 후 도면 열기
                    if (File.Exists(TStr))
                    {
                        EdApp.Documents.Open(TStr);
                    }
                    else
                    {
                        dynamic EdgeApplication = EdApp;
                        EdgeApplication.StartCommand("11604");
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
                Debug.Print(ex.Message);
                //Console.WriteLine(ex.Message);
            }
            finally
            {
                //OleMessageFilter.Revoke();
            }
        }
    }
}
