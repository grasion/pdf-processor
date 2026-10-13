using System;
using Velopack;

namespace PDFProcessor
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            // Process installation/update lifecycle hooks before creating any WPF windows.
            VelopackApp.Build().Run();
            var app = new App();
            app.InitializeComponent();
            app.Run();
        }
    }
}
