using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Spreadsheet {
    public class Extensions {
        private ScriptEngine _engine;
        private ScriptScope _scope;

        public Extensions() {
            InitializeScripts();
        }

        private void InitializeScripts() {
            _engine = Python.CreateEngine();
            _scope = _engine.Runtime.CreateScope();
        }

        public object Execute(string code) {
            Code = code;
            return _engine.Execute(Code, _scope);
        }

        public string Code { get; set; }
    }
}