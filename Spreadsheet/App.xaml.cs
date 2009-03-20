using System;
using System.Collections.Generic;
using System.Net;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Resources;

using Microsoft.Scripting.Silverlight;

namespace Spreadsheet {
    public partial class App : Application {

        public App() {
            this.Startup += this.Application_Startup;
            this.Exit += this.Application_Exit;
            this.UnhandledException += this.Application_UnhandledException;

            InitializeComponent();
        }

        private void Application_Startup(object sender, StartupEventArgs e) {
            DynamicApplication.LoadAssemblies(delegate() {
                this.RootVisual = new Page();
            });

            // download eggs
            if (HtmlPage.Document.QueryString.ContainsKey("test")) {
                var xap = new Uri("eggs.xap", UriKind.Relative);
                WebClient wcXap = new WebClient();
                wcXap.OpenReadCompleted += new OpenReadCompletedEventHandler(wcXap_OnOpenReadCompleted);
                wcXap.OpenReadAsync(xap);
            }
        }

        private void wcXap_OnOpenReadCompleted(object sender, OpenReadCompletedEventArgs e) {
            if ((e.Error == null) && (e.Cancelled == false)) {
                var testList = new Dictionary<string, List<string>>() {
                    {"tests", new List<string>() { "spreadsheet" }}
                };

                var xap = new StreamResourceInfo(e.Result, null);
                System.Reflection.Assembly asm = new AssemblyPart().Load(
                    Application.GetResourceStream(
                        xap, new Uri("Eggs.dll", UriKind.Relative)
                    ).Stream
                );
                asm.GetType("Eggs").GetMethod("Start").Invoke(null, new object[] {
                    (object) testList, (object) xap
                });
            }
        }

        private void Application_Exit(object sender, EventArgs e) {

        }
        private void Application_UnhandledException(object sender, ApplicationUnhandledExceptionEventArgs e) {
            // If the app is running outside of the debugger then report the exception using
            // the browser's exception mechanism. On IE this will display it a yellow alert 
            // icon in the status bar and Firefox will display a script error.
            if (!System.Diagnostics.Debugger.IsAttached) {

                // NOTE: This will allow the application to continue running after an exception has been thrown
                // but not handled. 
                // For production applications this error handling should be replaced with something that will 
                // report the error to the website and stop the application.
                e.Handled = true;
                Deployment.Current.Dispatcher.BeginInvoke(delegate { ReportErrorToDOM(e); });
            }
        }
        private void ReportErrorToDOM(ApplicationUnhandledExceptionEventArgs e) {
            try {
                string errorMsg = e.ExceptionObject.Message + e.ExceptionObject.StackTrace;
                errorMsg = errorMsg.Replace('"', '\'').Replace("\r\n", @"\n");

                System.Windows.Browser.HtmlPage.Window.Eval("throw new Error(\"Unhandled Error in Silverlight Application " + errorMsg + "\");");
            } catch (Exception) {
            }
        }
    }
}
