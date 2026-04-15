namespace Microsoft.Office.Interop.Excel
{
    public interface Application { Worksheet ActiveSheet { get; } }
    public interface Worksheet { Range get_Range(object Cell1, object Cell2 = null); Range this[object cell] { get; } }
    public interface Range { object Value2 { get; set; } }
}
namespace Microsoft.Office.Tools.Excel
{
    public class AddInBase
    {
        public event System.EventHandler Startup;
        public event System.EventHandler Shutdown;

        public Microsoft.Office.Interop.Excel.Application Application { get; set; }

        protected virtual void OnStartup() { Startup?.Invoke(this, System.EventArgs.Empty); }
        protected virtual void OnShutdown() { Shutdown?.Invoke(this, System.EventArgs.Empty); }
        protected virtual void Initialize() { }
        protected virtual void FinishInitialization() { }
        protected virtual void InitializeDataBindings() { }
    }

    public class Factory { }
}

namespace Microsoft.VisualStudio.Tools.Applications.Runtime
{
    public class StartupObjectAttribute : System.Attribute { public StartupObjectAttribute(int i) {} }
}

namespace Microsoft.Office.Tools
{
    public class CustomTaskPaneCollection : System.IDisposable
    {
        public void Dispose() { }
        public void BeginInit() { }
        public void EndInit() { }
    }
    public class SmartTagCollection : System.IDisposable
    {
        public void Dispose() { }
        public void BeginInit() { }
        public void EndInit() { }
    }

    namespace Ribbon
    {
        public class RibbonCollectionBase
        {
            public RibbonCollectionBase(RibbonFactory factory) {}
        }
        public class RibbonFactory {}
    }
}
