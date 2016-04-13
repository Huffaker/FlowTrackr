using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Lync.Model;

namespace FlowTrackr
{
    class Program
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;

        private static LowLevelKeyboardProc s_proc = KeyboardHookCallback;
        private static LowLevelKeyboardProc s_procMouse = MouseHookCallback;
        private static IntPtr s_hookId = IntPtr.Zero;
        private static IntPtr s_mouseHookId = IntPtr.Zero;

        private static IntPtr s_handle = IntPtr.Zero;

        private static LyncClient s_lyncClient;
        private static object s_lyncClient_lock = new object();

        private static ContactAvailability s_previousPresence;
        private static ContactAvailability s_currentPresence;

        private static object s_keyCountInPeriod_lock = new object();
        private static int s_keyCountInPeriod;
        private const int INTERVAL_TIME_MS = 30000;
        private static System.Threading.Timer s_timer;

        static void Main(string[] args)
        {
            Initialize();
            s_hookId = SetHook(s_proc);
            s_mouseHookId = SetHookMouse(s_procMouse);
            s_timer = new System.Threading.Timer(Timer_Elapsed, null, 0, INTERVAL_TIME_MS);

            Application.Run();
            //UnhookWindowsHookEx(s_hookId);
            //UnhookWindowsHookEx(s_mouseHookId);
            Thread.Sleep(-1);
        }

        static void Initialize()
        {
            s_handle = GetConsoleWindow();

            handler = new ConsoleEventDelegate(ConsoleEventCallback);
            SetConsoleCtrlHandler(handler, true);

            CreateLyncClient();

            ShowWindow(s_handle, SW_HIDE);
        }

        static void Deinitialize()
        {

            UnhookWindowsHookEx(s_hookId);
            UnhookWindowsHookEx(s_mouseHookId);
        }

        static LyncClient CreateLyncClient()
        {

            do
            {
                try
                {
                    s_lyncClient = null;
                    s_lyncClient = LyncClient.GetClient();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lync is not running, please start it. {ex}");
                    Thread.Sleep(5000);
                }
            } while (s_lyncClient == null);

            return s_lyncClient;
        }

        static bool ConsoleEventCallback(int eventType)
        {
            UnhookWindowsHookEx(s_hookId);
            UnhookWindowsHookEx(s_mouseHookId);
            SetPresence(ContactAvailability.Free);
            s_timer.Dispose();
            if (eventType == 2)
            {
                Console.WriteLine("Console window closing, death imminent");
            }
            return false;
        }
        static ConsoleEventDelegate handler;   // Keeps it from getting garbage collected
                                               // Pinvoke
        private delegate bool ConsoleEventDelegate(int eventType);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);


        private static void Timer_Elapsed(object state)
        {
            lock (s_keyCountInPeriod_lock)
            {
                SetPresence(GetPresence(s_keyCountInPeriod), GetMessage(s_keyCountInPeriod));
                s_keyCountInPeriod = 0;
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                lock (s_keyCountInPeriod_lock)
                {
                    s_keyCountInPeriod++;
                    Debug.WriteLine(s_keyCountInPeriod);
                }
                //var vkCode = Marshal.ReadInt32(lParam);
                //var code = (Keys)vkCode;
                //if (code == Keys.B)
                //{
                //    SetPresence(ContactAvailability.Busy);
                //}
                //else if (code == Keys.F)
                //{
                //    SetPresence(ContactAvailability.Free);
                //}
                //else if (code == Keys.D)
                //{
                //    SetPresence(ContactAvailability.DoNotDisturb);
                //}

                ////Console.WriteLine(code);
                //using (var sw = new StreamWriter(Application.StartupPath + "\\log.txt", true))
                //{
                //    sw.Write(code);
                //}
            }

            return CallNextHookEx(s_hookId, nCode, wParam, lParam);
        }

        private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_LBUTTONDOWN || wParam == (IntPtr)WM_RBUTTONDOWN))
            {
                lock (s_keyCountInPeriod_lock)
                {
                    s_keyCountInPeriod++;
                    Debug.WriteLine(s_keyCountInPeriod);

                }
            }

            return CallNextHookEx(s_mouseHookId, nCode, wParam, lParam);
        }

        private static void SetPresence(ContactAvailability contactAvailability, string message = null)
        {
            try
            {
                lock (s_lyncClient_lock)
                {
                    s_previousPresence = s_currentPresence;
                    s_currentPresence = contactAvailability;

                    //Check current Lync status
                    var lyncClientStatus =
                        (ContactAvailability)
                            s_lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);

                    if (s_previousPresence == lyncClientStatus)
                    {
                        //Updating
                        var newInformation = new Dictionary<PublishableContactInformationType, object>()
                        {
                            {PublishableContactInformationType.Availability, contactAvailability},
                            {PublishableContactInformationType.PersonalNote, message ?? string.Empty}
                        };

                        s_lyncClient.Self.BeginPublishContactInformation(newInformation, null, null);
                    }
                    else
                    {
                        //Lync status was changed by user, set state to free, we will wait until the user goes back to free before continuing
                        s_previousPresence = ContactAvailability.Free;
                        s_currentPresence = ContactAvailability.Free;
                    }

                }
            }
            catch (Exception ex)
            {
                try
                {
                    ShowWindow(s_handle, SW_SHOW);
                    Console.WriteLine(ex);
                }
                catch
                {
                    // ignored
                }
                finally
                {
                    //Deinitialize();
                    Initialize();
                }
            }
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            {
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        private static IntPtr SetHookMouse(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            {
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        private static ContactAvailability GetPresence(int count)
        {
            if (count < 100)
            {
                return ContactAvailability.Free;
            }
            else if (count < 200)
            {
                return ContactAvailability.Busy;
            }
            else if (count >= 200)
            {
                return ContactAvailability.DoNotDisturb;
            }

            throw new Exception();
        }

        private static string GetMessage(int count)
        {
            if (count < 100)
            {
                if (s_previousPresence == ContactAvailability.DoNotDisturb || s_previousPresence == ContactAvailability.Busy)
                {
                    return "Lost the flow";
                }
                return string.Empty;
            }
            else if (count < 200)
            {
                if (s_previousPresence == ContactAvailability.DoNotDisturb)
                {
                    return "Losing the flow";
                }
                else if (s_previousPresence == ContactAvailability.Free || s_previousPresence == ContactAvailability.None)
                {
                    return "Getting into the flow";
                }

                return "Weak flow";
            }
            else if (count < 300)
            {
                return "In the flow";
            }
            else if (count >= 300)
            {
                return "Letting the code flow";
            }

            throw new Exception();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
    }
}
