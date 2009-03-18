﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Spreadsheet {
    public abstract class ViewModelBase : INotifyPropertyChanged, IDisposable {
        protected ViewModelBase() { }

        /// <summary>
        /// Returns the user-friendly name of this object.
        /// Child classes can set this property to a new value,
        /// or override it to determine the value on-demand.
        /// </summary>
        internal virtual string DisplayName { get; private set; }

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        protected virtual void OnPropertyChanged(string propertyName) {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null) {
                var e = new PropertyChangedEventArgs(propertyName);
                handler(this, e);
            }
        }

        #endregion // INotifyPropertyChanged Members

        #region IDisposable Members

        /// <summary>
        /// Invoked when this object is being removed from the application
        /// and will be subject to garbage collection.
        /// </summary>
        public void Dispose() {
            this.OnDispose();
        }

        /// <summary>
        /// Child classes can override this method to perform 
        /// clean-up logic, such as removing event handlers.
        /// </summary>
        protected virtual void OnDispose() {
        }

#if DEBUG
        /// <summary>
        /// Useful for ensuring that ViewModel objects are properly garbage collected.
        /// </summary>
        ~ViewModelBase() {
            string msg = string.Format("{0} ({1}) ({2}) Finalized", this.GetType().Name, this.DisplayName, this.GetHashCode());
            System.Diagnostics.Debug.WriteLine(msg);
        }
#endif

        #endregion // IDisposable Members
    }

    public class RowViewModelBase : ViewModelBase {
        internal SpreadsheetViewModel ViewModel { get; private set; }
        private string _rowNumber;

        public RowViewModelBase(SpreadsheetViewModel viewModel, int rowNumber) {
            ViewModel = viewModel;
            _rowNumber = rowNumber.ToString();
        }

        public string GetCell(string column) {
            if (ViewModel.Editable)
                return GetExpression(column);
            else
                return ViewModel.Model.GetCell(column + _rowNumber);
        }

        public string GetExpression(string column) {
            return ViewModel.Model.GetExpression(column + _rowNumber);
        }

        public void SetCell(string column, string value) {
            base.OnPropertyChanged(column);
            ViewModel.Model.SetCell(column + _rowNumber, value);
        }

        public static Type RowType;

        private static void CreateProperty(string name, TypeBuilder tb, MethodInfo bgcmi, MethodInfo bscmi) {
            // define getter
            var gmb = tb.DefineMethod("get_" + name, 
                MethodAttributes.SpecialName | MethodAttributes.HideBySig | MethodAttributes.Public,
                typeof(string),
                null);

            var g = gmb.GetILGenerator();
            g.Emit(OpCodes.Ldarg_0);
            g.Emit(OpCodes.Ldstr, name);
            g.Emit(OpCodes.Call, bgcmi);
            g.Emit(OpCodes.Ret);

            // define setter
            var smb = tb.DefineMethod("set_" + name,
                MethodAttributes.SpecialName | MethodAttributes.HideBySig | MethodAttributes.Public,
                null,
                new Type[] { typeof(string) });

            g = smb.GetILGenerator();
            g.Emit(OpCodes.Ldarg_0);
            g.Emit(OpCodes.Ldstr, name);
            g.Emit(OpCodes.Ldarg_1);
            g.Emit(OpCodes.Call, bscmi);
            g.Emit(OpCodes.Ret);
            
            // define properties
            var pb = tb.DefineProperty(name, PropertyAttributes.None, typeof(string), null);
            pb.SetGetMethod(gmb);
            pb.SetSetMethod(smb);
        }

        private static Type CreateType(ModuleBuilder mb, int columns) {
            var tb = mb.DefineType("Test", 
                TypeAttributes.Public | TypeAttributes.BeforeFieldInit | TypeAttributes.AnsiClass | TypeAttributes.AutoClass,
                typeof(RowViewModelBase));

            // base constructor reference
            var bc = typeof(RowViewModelBase).GetConstructor(new Type[] { typeof(SpreadsheetViewModel), typeof(int) });

            // define constructor
            var cb = tb.DefineConstructor(MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName,
                CallingConventions.Standard, 
                new Type[] { typeof(SpreadsheetViewModel), typeof(int) });
            var g = cb.GetILGenerator();
            g.Emit(OpCodes.Ldarg_0);
            g.Emit(OpCodes.Ldarg_1);
            g.Emit(OpCodes.Ldarg_2);
            g.Emit(OpCodes.Call, bc);
            g.Emit(OpCodes.Ret);

            // base GetCell methodinfo
            var bgcmi = typeof(RowViewModelBase).GetMethod("GetCell");
            var bscmi = typeof(RowViewModelBase).GetMethod("SetCell");

            for (int i = 0; i < columns; i++) {
                CreateProperty(CellParser.GenerateColumnName(i), tb, bgcmi, bscmi);
            }
            
            return tb.CreateType();
        }

        public static Type Initialize(ModuleBuilder mb, int columns) {
            RowType = CreateType(mb, columns);
            return RowType;
        }

        public static RowViewModelBase Create(ModuleBuilder mb, SpreadsheetViewModel viewModel, int rowNumber, int columns) {
            if (RowType == null)
                RowType = CreateType(mb, columns);

            return (RowViewModelBase)Activator.CreateInstance(RowType, viewModel, rowNumber);
        }
    }

    // TODO: build a delegate system that will invoke Add and indexer against OC<T>
    public class SpreadsheetViewModel : ViewModelBase {
        private object _rows;
        private SpreadsheetModel _model;
        private MethodInfo _getItem;

        public delegate void AddDelegate(object collection, object obj);

        public SpreadsheetViewModel(ModuleBuilder mb, int rows, int cols) {
            Type rowType = RowViewModelBase.Initialize(mb, cols);
            Type gt = typeof(ObservableCollection<>);
            Type cgt = gt.MakeGenericType(rowType);

            MethodInfo add = cgt.GetMethod("Add");
            _getItem = cgt.GetMethod("get_Item");

            _rows = Activator.CreateInstance(cgt);
            _model = new SpreadsheetModel();

            DynamicMethod dm = new DynamicMethod("Add", null, new Type[] { typeof(object), typeof(object) });
            var g = dm.GetILGenerator();
            g.Emit(OpCodes.Ldarg_0);
            g.Emit(OpCodes.Castclass, cgt);
            g.Emit(OpCodes.Ldarg_1);
            g.Emit(OpCodes.Castclass, rowType);
            g.Emit(OpCodes.Call, add);
            g.Emit(OpCodes.Ret);

            AddDelegate m = (AddDelegate)dm.CreateDelegate(typeof(AddDelegate));
            
            for (int i = 0; i < rows; i++) {
                RowViewModelBase row = RowViewModelBase.Create(mb, this, i + 1, cols);
                m(_rows, row);
            }
        }

        public IEnumerable DataSource { get { return (IEnumerable)_rows; } }
        public bool Editable { get; set; }
        public SpreadsheetModel Model { get { return _model; } }

        public string GetCell(string cell) {
            // TODO: build dynamic method to invoke this 
            var coord = CellParser.ParseCellName(cell);
            RowViewModelBase row = (RowViewModelBase)_getItem.Invoke(_rows, new object[] { coord.Row });
            return row.GetCell(coord.Col);
        }

        public void SetCell(string cell, string value) {
            // TODO: build dynamic method to invoke this
            var coord = CellParser.ParseCellName(cell);
            RowViewModelBase row = (RowViewModelBase)_getItem.Invoke(_rows, new object[] { coord.Row });
            row.SetCell(coord.Col, value);
        }
    }
}