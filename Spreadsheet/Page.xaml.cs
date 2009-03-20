using System;
using System.IO;
using System.Net;
using System.Reflection;
using System.Reflection.Emit;
using System.Windows.Browser;
using System.Windows.Controls;
using System.Windows.Input;

namespace Spreadsheet {
    public partial class Page : UserControl {
        private SpreadsheetViewModel _vm;
        private Extensions _extensions;

        public Page() {
            InitializeComponent();
            InitializeSpreadsheet();
        }

        public SpreadsheetViewModel Model { get { return _vm; } }

        private ModuleBuilder CreateDynamicModule() {
            var name = new AssemblyName("test");
            var ab = AppDomain.CurrentDomain.DefineDynamicAssembly(name, AssemblyBuilderAccess.Run);
            return ab.DefineDynamicModule("test");
        }

        private void InitializeSpreadsheet() {
            LoadFunctions();
            _extensions = new Extensions();
            _vm = new SpreadsheetViewModel(_extensions, CreateDynamicModule(), 15, 52);
            Spreadsheet.ItemsSource = _vm.DataSource;
        }

        private void Spreadsheet_LoadingRow(object sender, DataGridRowEventArgs e) {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Spreadsheet_KeyUp(object sender, System.Windows.Input.KeyEventArgs e) {
            if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                return;

            Spreadsheet.BeginEdit(e);
        }

        private void Spreadsheet_BeginningEdit(object sender, DataGridBeginningEditEventArgs e) {
            _vm.Editable = true;
        }

        private void Spreadsheet_CellEditEnded(object sender, DataGridCellEditEndedEventArgs e) {
            _vm.Editable = false;
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e) {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Enter)
                _extensions.Execute(CodeTextBox.Text);
        }

        #region Save Button

        private void SaveButton_Click(object sender, System.Windows.RoutedEventArgs e) {
            SaveButton.Content = "Saving ...";

            var uri = new Uri(HtmlPage.Document.DocumentUri, "/Script/save");

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "POST"; 
            request.ContentType = "application/x-www-form-urlencoded"; 
            request.BeginGetRequestStream(new AsyncCallback(SaveFunctions_RequestProceed), request);
        }

        void SaveFunctions_RequestProceed(IAsyncResult asyncResult) {
            HttpWebRequest request = asyncResult.AsyncState as HttpWebRequest;
            Stream postData = request.EndGetRequestStream(asyncResult);

            this.Dispatcher.BeginInvoke(delegate() {
                StreamWriter writer = new StreamWriter(postData);
                string code = Uri.EscapeDataString(CodeTextBox.Text);
                writer.Write(string.Format("code={0}", code));
                writer.Close();

                request.BeginGetResponse(new AsyncCallback(SaveFunctions_ResponseProceed), request);
            });
        }

        void SaveFunctions_ResponseProceed(IAsyncResult asyncResult) {
            HttpWebRequest request = asyncResult.AsyncState as HttpWebRequest;
            StreamReader reader;
            try {
                HttpWebResponse response = (HttpWebResponse)request.EndGetResponse(asyncResult);
                reader = new StreamReader(response.GetResponseStream());
                this.Dispatcher.BeginInvoke(delegate() {
                    string result = reader.ReadToEnd();
                    if (result != "True") {
                        HtmlPage.Window.Alert("Error, please try again");
                    }
                });
            } catch (Exception e) {
                this.Dispatcher.BeginInvoke(delegate() {
                    HtmlPage.Window.Alert("Error, please try again");
                });
            }
            this.Dispatcher.BeginInvoke(delegate() {
                SaveButton.Content = "Save";
            });
        }

        #endregion

        #region Load Button

        private void LoadButton_Click(object sender, System.Windows.RoutedEventArgs e) {
            LoadFunctions();
        }

        private void LoadFunctions() {
            CodeTextBox.Text = "Loading ...";
            WebClient wc = new WebClient(); 
            string sRequest = "/Script";
            wc.DownloadStringCompleted += new DownloadStringCompletedEventHandler(LoadFunctions_Completed);
            wc.DownloadStringAsync(new Uri(sRequest, UriKind.Relative));
        }

        void LoadFunctions_Completed(object sender, DownloadStringCompletedEventArgs e) {
            _extensions.Execute(e.Result);
            CodeTextBox.Text = e.Result;
        }

        #endregion
    }
}