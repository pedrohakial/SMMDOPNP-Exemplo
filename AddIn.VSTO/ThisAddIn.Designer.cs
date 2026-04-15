namespace AddIn.VSTO {

    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(1)]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.Excel.AddInBase {

        internal Microsoft.Office.Interop.Excel.Application Application;

        public ThisAddIn() : base() { }

        protected override void Initialize() {
            base.Initialize();
            this.Application = new MockApplication();
            this.InternalStartup();
        }

        private class MockApplication : Microsoft.Office.Interop.Excel.Application {
            public Microsoft.Office.Interop.Excel.Worksheet ActiveSheet => new MockWorksheet();
        }

        private class MockWorksheet : Microsoft.Office.Interop.Excel.Worksheet {
            public Microsoft.Office.Interop.Excel.Range get_Range(object Cell1, object Cell2 = null) => new MockRange();
            public Microsoft.Office.Interop.Excel.Range this[object cell] => new MockRange();
        }

        private class MockRange : Microsoft.Office.Interop.Excel.Range {
            public object Value2 { get; set; }
        }
    }
}
