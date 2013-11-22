using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Tools.WindowsInstallerXml.Bootstrapper;
using System.Windows.Threading;

namespace fy24hSetupBootstrapper
{
    public class CRMCustomBootstrapper : BootstrapperApplication
    {
        // global dispatcher
        static public Dispatcher BootstrapperDispatcher { get; private set; }

        protected override void Run()
        {
            this.Engine.Log(LogLevel.Verbose, "Launching custom UX");
            BootstrapperDispatcher = Dispatcher.CurrentDispatcher;

            MainViewModel viewModel = new MainViewModel(this);
            viewModel.Bootstrapper.Engine.Detect();

            MainView view = new MainView();
            view.DataContext = viewModel;
            view.Closed += (sender, e) => BootstrapperDispatcher.InvokeShutdown();
            view.Show();

            Dispatcher.Run();

            this.Engine.Quit(0);
        }

    }
}
