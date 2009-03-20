using System;
using System.Reflection;
using System.Reflection.Emit;
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

        private void InitializeSpreadsheet() {
            var name = new AssemblyName("test");
            var ab = AppDomain.CurrentDomain.DefineDynamicAssembly(name, AssemblyBuilderAccess.Run);
            var mb = ab.DefineDynamicModule("test");
            _extensions = new Extensions();
            CodeTextBox.Text = _extensions.Code.TrimStart();
            _vm = new SpreadsheetViewModel(_extensions, mb, 15, 52);
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

        private void TextBox_KeyUp(object sender, KeyEventArgs e) {
        }

        bool _isCtrl = false;

        private void TextBox_KeyDown(object sender, KeyEventArgs e) {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Enter)
                _extensions.Execute(CodeTextBox.Text);
        }
    }
}