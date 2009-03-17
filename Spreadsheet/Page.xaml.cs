using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Scripting.Hosting;
using IronPython.Hosting;
using System.Windows;

namespace Spreadsheet {
    public partial class Page : UserControl {
        private SpreadsheetViewModel _vm;

        public Page() {
            InitializeComponent();
            InitializeSpreadsheet();
        }

        public SpreadsheetViewModel Model { get { return _vm; } }

        private void InitializeSpreadsheet() {
            _vm = new SpreadsheetViewModel(15, 4);
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
            // tell model that we are entering edit mode
            //_data.Editable = true;
        }

        private void Spreadsheet_CellEditEnded(object sender, DataGridCellEditEndedEventArgs e) {
            //_data.Editable = false;
        }
    }
}