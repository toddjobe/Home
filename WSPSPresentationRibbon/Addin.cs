using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using Extensibility;

using PowerPoint = NetOffice.PowerPointApi;
using Office = NetOffice.OfficeApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace WSPS
{
    [GuidAttribute("f88eaf02-ce89-4178-af38-38345d0b563d"), ProgId("WSPSPPointAddin.RibbonAddin"), ComVisible(true)]
    public class Addin : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private static readonly string _addinOfficeRegistryKey = "Software\\Microsoft\\Office\\PowerPoint\\AddIns\\";
        private static readonly string _progId = "WSPSPPointAddin.RibbonAddin";
        private static readonly string _addinFriendlyName = "WSPS PowerPoint Addin";
        private static readonly string _addinDescription = "WSPS PowerPoint Addin for recording sermons with custom Ribbon UI.";
        private static Thread _recThread;
        PowerPoint.Application _powerApplication;

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _powerApplication = new PowerPoint.Application(null, Application);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _powerApplication)
                    _powerApplication.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {

        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {

        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {

        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return ReadString("RibbonUI.xml");
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured in GetCustomUI." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
        }

        #endregion

        #region Ribbon Gui Trigger

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnStartSlideShow":
                        // Set the event listener for the beginning of the show
                        _powerApplication.SlideShowBeginEvent += new PowerPoint.Application_SlideShowBeginEventHandler(powerApplication_SlideShowBeginEvent);

                        // Set the event listener for the end of the show
                        _powerApplication.SlideShowEndEvent += new PowerPoint.Application_SlideShowEndEventHandler(powerApplication_SlideShowEndEvent);

                        // Start the Slide show
                        _powerApplication.ActiveWindow.Presentation.SlideShowSettings.Run();
                    
                        break;
                    case "customButton2":
                        MessageBox.Show("This is the second sample button.", _progId);
                        break;
                    default:
                        MessageBox.Show("Unkown Control Id: " + control.Id, _progId);
                        break;
                }
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Addin));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

                // add powerpoint addin key
                Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _progId);
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _progId, true);
                rk.SetValue("LoadBehavior", Convert.ToInt32(3));
                rk.SetValue("FriendlyName", _addinFriendlyName);
                rk.SetValue("Description", _addinDescription);
                rk.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _progId, false);
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Private Helper

        /// <summary>
        /// reads text from ressource
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static string ReadString(string fileName)
        {
            Assembly assembly = typeof(Addin).Assembly;
            System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(assembly.GetName().Name + "." + fileName);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource File."));

            string text = textStreamReader.ReadToEnd();
            ressourceStream.Close();
            textStreamReader.Close();
            return text;
        }


        /// <summary>
        /// Executes a shell command synchronously.
        /// </summary>
        /// <param name="command">string command</param>
        /// <returns>string, as output of the command.</returns>
        public void ExecuteCommandSync(object command)
        {
            try
            {
                // create the ProcessStartInfo using "cmd" as the program to be run,
                // and "/c " as the parameters.
                // Incidentally, /c tells cmd that we want it to execute the command that follows,
                // and then exit.
                System.Diagnostics.ProcessStartInfo procStartInfo =
                    new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);

                // The following commands are needed to redirect the standard output.
                // This means that it will be redirected to the Process.StandardOutput StreamReader.
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = true;
                // Now we create a process, assign its ProcessStartInfo and start it
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
            }
            catch (Exception objException)
            {
                // Log the exception
            }
        }

        /// <summary>
        /// Asychronous shell commands
        /// </summary>
        /// <param name="command">string command.</param>
        private Thread ExecuteCommandAsync(string command)
        {
            Thread objThread;
            try
            {
                //Asynchronously start the Thread to process the Execute command request.
                objThread = new Thread(new ParameterizedThreadStart(ExecuteCommandSync));
                //Make the thread as background thread.
                objThread.IsBackground = true;
                //Set the Priority of the thread.
                objThread.Priority = ThreadPriority.AboveNormal;
                //Start the thread.
                objThread.Start(command);
            }
            catch (ThreadStartException objException)
            {
                // Log the exception
                objThread = null;
            }
            catch (ThreadAbortException objException)
            {
                // Log the exception
                objThread = null;
            }
            catch (Exception objException)
            {
                // Log the exception
                objThread = null;
            }
            return objThread;
        }
        #endregion

        #region Event Handlers

        /// <summary>
        /// Slide Show Begin Event handler
        /// </summary>
        /// <param name="command">PowerPoint.Presentation pres.</param>
        private void powerApplication_SlideShowBeginEvent(PowerPoint.SlideShowWindow ssw)
        {
            bool test = true;
            if (test)
            {
                MessageBox.Show("Called SlideshowBeginEvent");
            }
            else
            {

                // Start the screencast recorder (depends on ffmpeg)
                // Could use its API which is libavformat, but I'm lazy right now
                /* pipe the screen as mpeg4 to a file or mux it
                * Not sure exactly how to 'quit' the process right now.  Normally, you'd type 'q' on terminal.
                * ffmpeg -f dshow  -i video="UScreenCapture"  -r 30 -vcodec mpeg4 -q 12 -f %temp%/sermon.mp4
                * This command below uses UDP for streaming which means packets could be lost or bandwidth could be gobbled up.
                * http://nerdlogger.com/2011/11/03/stream-your-windows-desktop-using-ffmpeg/
                ffmpeg -f dshow  -i video="UScreenCapture"  -r 30 -vcodec mpeg4 -q 12 -f mpegts udp://aaa.bbb.ccc.ddd:6666?pkt_size=188?buffer_size=65535
                */

                _recThread = ExecuteCommandAsync("ffmpeg -f dshow  -i video=\"UScreenCapture\"  -r 30 -vcodec mpeg4 -q 12 -f %temp%/sermon.mp4");
            }

        }

        /// <summary>
        /// Slide Show End Event handler
        /// </summary>
        /// <param name="command">PowerPoint.Presentation pres.</param>
        private void powerApplication_SlideShowEndEvent(PowerPoint.Presentation pres)
        {
            bool test = true;
            if (test)
            {
                MessageBox.Show("Called SlideshowEndEvent");
            }
            else
            {

                // quit recording video
                // Not sure this will work.  In fact, I'm almost sure it won't.
                _recThread.Start("q");

                // dispose of the slide show event listeners.
                // Not sure this will work nested.
                _powerApplication.SlideShowBeginEvent -= powerApplication_SlideShowBeginEvent;
                _powerApplication.SlideShowEndEvent += powerApplication_SlideShowEndEvent;
            }
        }

        #endregion

    }


}
