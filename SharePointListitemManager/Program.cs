using System;
using System.Windows.Forms;
using Autofac;

namespace SharePointListitemManager
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            var builder = new ContainerBuilder();

            builder.RegisterType<Form1>()
                .As<Form1>()
                .SingleInstance();
            builder.RegisterType<SharepointService>()
                .As<ISharepointService>();
            builder.RegisterType<ExcelService>()
                .As<IExcelService>();

            var container = builder.Build();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(container.Resolve<Form1>());
        }
    }
}
