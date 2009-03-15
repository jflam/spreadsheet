using System;
using System.Diagnostics;
using System.Text;

public enum Tokens {
    None,
    Add,
    Subtract,
    Multiply,
    Divide,
    Equals,
    Number,
    Function,
    CellReference,
    Whitespace,
}

public class Tokenizer {
    private string _buffer;
    private int _pos;
    private object _value;
    private Tokens _token;

    public Tokenizer(string expression) {
        _buffer = expression + null;
        _pos = 0;
    }

    private char Read() {
        return _buffer[_pos++];
    }

    private char ReadCellReference(char c) {
        _token = Tokens.CellReference;
        var start = _pos - 1;
        while (true) {
            if (_pos == _buffer.Length || !Char.IsLetterOrDigit(c)) {
                _value = _buffer.Substring(start, _pos - start - 1);
                return c;
            }
            c = Read();
        }
    }

    private void ReadFunction(char c) {
        _token = Tokens.Function;
        // Rest of buffer is assumed to be function expression
        _value = _buffer.Substring(_pos);
        _pos = _buffer.Length;
    }

    private char ReadNumber(char c) {
        _token = Tokens.Number;
        var start = _pos - 1;
        while (true) {
            if (_pos == _buffer.Length || c == '.' || !Char.IsDigit(c)) {
                _value = Double.Parse(_buffer.Substring(start, _pos - start));
                return c;
            }
            c = Read();
        }
    }

    public bool ReadNextToken() {
        char c = Read();
        while (true) {
            switch (c) {
                case '\t':
                case ' ':
                    _token = Tokens.Whitespace;
                    break;
                case '0':
                case '1':
                case '2':
                case '3':
                case '4':
                case '5':
                case '6':
                case '7':
                case '8':
                case '9':
                    c = ReadNumber(c);
                    break;
                case '+':
                    _token = Tokens.Add;
                    break;
                case '-':
                    _token = Tokens.Subtract;
                    break;
                case '*':
                    _token = Tokens.Multiply;
                    break;
                case '/':
                    _token = Tokens.Divide;
                    break;
                case '@':
                    ReadFunction(c);
                    return false;
                default:
                    if (Char.IsLetter(c))
                        c = ReadCellReference(c);
                    break;
            }
            if (_pos >= _buffer.Length)
                return false;
            if (_token != Tokens.Whitespace)
                return true;
            c = Read();
        }
    }

    public object CurrentValue {
        get { return _value; }
    }

    public Tokens CurrentToken {
        get { return _token; }
    }
}