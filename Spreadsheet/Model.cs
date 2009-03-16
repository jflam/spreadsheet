using System;
using System.Collections;
using System.Collections.ObjectModel;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using Spreadsheet;
using System.Windows;

public class Model {
    private ObservableCollection<Row> _rows;
    private ScriptEngine _engine;
    private ScriptScope _scope;

    public Model(int rows, int cols) {
        _rows = new ObservableCollection<Row>();
        for (int i = 0; i < rows; i++) {
            _rows.Add(Row.Create(this, cols));
        }
        _engine = Python.CreateEngine();
        _scope = _engine.Runtime.CreateScope();
        var sum = @"
def sum2(*args):
    sum = 0
    for arg in args:
        sum += arg
    return sum
";
        _engine.Execute(sum, _scope);
    }

    public IEnumerable DataSource { get { return _rows; } }
    public bool Editable { get; set; }

    private void Crack(string cellDescriptor, out int row, out int col) {
        char colName = cellDescriptor[0];
        col = colName - 'A';
        row = int.Parse(cellDescriptor.Substring(1)) - 1;
    }

    public string GetCell(string cellDescriptor) {
        int row, col;
        Crack(cellDescriptor, out row, out col);
        return _rows[row].GetCell(col);
    }

    public void SetCell(string cellDescriptor, string value) {
        int row, col;
        Crack(cellDescriptor, out row, out col);
        _rows[row].SetCell(col, value);
    }

    public string Calc(string expression) {
        if (String.IsNullOrEmpty(expression))
            return String.Empty;

        var tokenizer = new Tokenizer(expression);
        double acc = 0;
        double currentValue;
        Tokens state = Tokens.None;
        while (true) {
            bool resume = tokenizer.ReadNextToken();
            Tokens currentToken = tokenizer.CurrentToken;
            switch (currentToken) {
                case Tokens.Number:
                case Tokens.CellReference:
                    currentValue = currentToken == Tokens.Number ? (double)tokenizer.CurrentValue
                                                                 : Double.Parse(GetCell((string)tokenizer.CurrentValue));
                    switch (state) {
                        case Tokens.None:
                            acc = currentValue;
                            state = Tokens.Number;
                            break;
                        case Tokens.Add:
                            acc = acc + currentValue;
                            break;
                        case Tokens.Subtract:
                            acc = acc - currentValue;
                            break;
                        case Tokens.Multiply:
                            acc = acc * currentValue;
                            break;
                        case Tokens.Divide:
                            acc = acc / currentValue;
                            break;
                    }
                    state = Tokens.Number;
                    break;
                case Tokens.Add:
                case Tokens.Subtract:
                case Tokens.Multiply:
                case Tokens.Divide:
                    if (state == Tokens.Number) {
                        state = tokenizer.CurrentToken;
                    } else {
                        throw new ArgumentException("expected value but found an operator");
                    }
                    break;
                case Tokens.Function:
                    return _engine.Execute(tokenizer.CurrentValue.ToString(), _scope).ToString();
            }
            if (!resume)
                break;
        }
        return acc.ToString();
    }
}


public class Row : DependencyObject {
    private Model _model;
    private string[] _cells;
    private static Type _rowType;

    public static Row Create(Model model, int cols) {
        if (_rowType == null) {
            // TODO: create a new subtype of Row on demand if _rowType is not defined yet
        }
        return new Row(model, cols);
    }

    private Row(Model model, int cols) {
        _model = model;
        _cells = new string[26]; // TODO: parameterize with cols
    }

    internal Model Model { get { return _model; } }

    public string GetCell(int col) {
        return _model.Editable ? _cells[col] : _model.Calc(_cells[col]);
    }

    public void SetCell(int col, string value) {
        _cells[col] = value + Char.MinValue;
    }

    #region Embarrassing
    static Row() {
        AProperty = DependencyProperty.Register("A", typeof(string), typeof(Row), new PropertyMetadata(String.Empty));
    }

    public static readonly DependencyProperty AProperty;

    // TODO: gen a subclass of Row that contains dynamically generated properties based on # cols desired
    public string A { get { 
        return (string)GetValue(AProperty); } 
        set { SetValue(AProperty, value); } }
    //public string A { get { return GetCell(0); } set { SetCell(0, value); } }
    public string B { 
        get { return GetCell(1); }
        set { SetCell(1, value); } }
    public string C { get { return GetCell(2); } set { SetCell(2, value); } }
    public string D { get { return GetCell(3); } set { SetCell(3, value); } }
    public string E { get { return GetCell(4); } set { SetCell(4, value); } }
    public string F { get { return GetCell(5); } set { SetCell(5, value); } }
    public string G { get { return GetCell(6); } set { SetCell(6, value); } }
    public string H { get { return GetCell(7); } set { SetCell(7, value); } }
    public string I { get { return GetCell(8); } set { SetCell(8, value); } }
    public string J { get { return GetCell(9); } set { SetCell(9, value); } }
    public string K { get { return GetCell(10); } set { SetCell(10, value); } }
    public string L { get { return GetCell(11); } set { SetCell(11, value); } }
    public string M { get { return GetCell(12); } set { SetCell(12, value); } }
    public string N { get { return GetCell(13); } set { SetCell(13, value); } }
    public string O { get { return GetCell(14); } set { SetCell(14, value); } }
    public string P { get { return GetCell(15); } set { SetCell(15, value); } }
    public string Q { get { return GetCell(16); } set { SetCell(16, value); } }
    public string R { get { return GetCell(17); } set { SetCell(17, value); } }
    public string S { get { return GetCell(18); } set { SetCell(18, value); } }
    public string T { get { return GetCell(19); } set { SetCell(19, value); } }
    public string U { get { return GetCell(20); } set { SetCell(20, value); } }
    public string V { get { return GetCell(21); } set { SetCell(21, value); } }
    public string W { get { return GetCell(22); } set { SetCell(22, value); } }
    public string X { get { return GetCell(23); } set { SetCell(23, value); } }
    public string Y { get { return GetCell(24); } set { SetCell(24, value); } }
    public string Z { get { return GetCell(25); } set { SetCell(25, value); } }
    #endregion
}