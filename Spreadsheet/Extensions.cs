using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Spreadsheet {
    public class Extensions {
        private ScriptEngine _engine;
        private ScriptScope _scope;
        public const string DEFAULT_CODE = @"
def sum(*args):
    sum = 0
    for arg in args:
        sum += arg
    return sum
";

        public Extensions() {
            InitializeScripts();
        }

        private void InitializeScripts() {
            _engine = Python.CreateEngine();
            _scope = _engine.Runtime.CreateScope();
            Code = Load();
            _engine.Execute(Code, _scope);
        }

        public object Execute(string code) {
            Code = code;
            return _engine.Execute(Code, _scope);
        }

        public string Load() {
            return DEFAULT_CODE;
        }

        public void Save(string code) {
        }

        public string Code { get; set; }
    }
}