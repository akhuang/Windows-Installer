using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Tools.WindowsInstallerXml.Bootstrapper;
using CustomBootstrapper.Core;
using System.Windows.Input;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Interop;

namespace fy24hSetupBootstrapper
{
    public class MainViewModel
    {
        public MainViewModel()
        {
            InitialComponentStatus();
        }

        public BootstrapperApplication Bootstrapper { get; private set; }
        public MainViewModel(BootstrapperApplication bootstrapper)
            : this()
        {
            this.Bootstrapper = bootstrapper;
            this.Bootstrapper.DetectPackageComplete += this.OnDetectPackageComplete;
            this.Bootstrapper.PlanComplete += this.OnPlanComplete;
            this.Bootstrapper.ApplyComplete += this.OnApplyComplete;
        }

        private void OnDetectPackageComplete(object sender, DetectPackageCompleteEventArgs e)
        {
            if (e.PackageId == "NetFx40Full")
            {
                Net40Present = e.State == PackageState.Present;

                //else if (e.State == PackageState.Present)
                //    UninstallEnabled = true;
            }
        }

        /// <summary>
        /// Method that gets invoked when the Bootstrapper PlanComplete event is fired.
        /// If the planning was successful, it instructs the Bootstrapper Engine to 
        /// install the packages.
        /// </summary>
        private void OnPlanComplete(object sender, PlanCompleteEventArgs e)
        {
            if (e.Status >= 0)
                Bootstrapper.Engine.Apply(System.IntPtr.Zero);

            //if (e.Status >= 0)
            //{
            //    MainView mv = new MainView();
            //    var mainWindowHandle = new WindowInteropHelper(mv).EnsureHandle();
            //    Bootstrapper.Engine.Apply(mainWindowHandle);
            //}
        }


        /// <summary>
        /// Method that gets invoked when the Bootstrapper ApplyComplete event is fired.
        /// This is called after a bundle installation has completed. Make sure we updated the view.
        /// </summary>
        private void OnApplyComplete(object sender, ApplyCompleteEventArgs e)
        {
            InstallEnabled = false;
            //UninstallEnabled = false;
        }

        private void InitialComponentStatus()
        {
            IIS6Present = InternetInformationServicesDetection.IsInstalled(ConstantEnum.InternetInformationServicesVersion.IIS6);
            IIS7Present = InternetInformationServicesDetection.IsInstalled(ConstantEnum.InternetInformationServicesVersion.IIS7);

            Sql2008Present = SqlServerDetection.IsInstalled(ConstantEnum.SqlServerVersion.Sql2008);
            Sql2012Present = SqlServerDetection.IsInstalled(ConstantEnum.SqlServerVersion.Sql2012);
        }

        private void InstallExecute()
        {
            Bootstrapper.Engine.Log(LogLevel.Verbose, "Lauch InstallExecute()");

            Bootstrapper.Engine.Plan(LaunchAction.Install);
        }

        //private void UninstallExecute()
        //{
        //    Bootstrapper.Engine.Plan(LaunchAction.Uninstall);
        //}

        private void ExitExecute()
        {
            CRMCustomBootstrapper.BootstrapperDispatcher.InvokeShutdown();
        }

        #region Commands

        private ICommand installCommand;
        public ICommand InstallCommand
        {
            get
            {
                //if (installCommand == null)
                //    installCommand = new RelayCommand(() => InstallExecute(), () => InstallEnabled == true);
                Bootstrapper.Engine.Log(LogLevel.Verbose, "InstallEnabled::" + InstallEnabled);
                return installCommand ?? (new CommandHandler(() => InstallExecute(), () => InstallEnabled == true));

                //return installCommand;
            }
        }

        //private ICommand uninstallCommand;
        //public ICommand UninstallCommand
        //{
        //    get
        //    {
        //        if (uninstallCommand == null)
        //            uninstallCommand = new RelayCommand(() => UninstallExecute(), () => UninstallEnabled == true);

        //        return uninstallCommand;
        //    }
        //}

        private ICommand exitCommand;
        public ICommand ExitCommand
        {
            get
            {

                return exitCommand ?? new CommandHandler(() => ExitExecute());
            }
        }

        #endregion //RelayCommands

        #region Properties
        private bool _installenabled;
        public bool InstallEnabled
        {
            get
            {
                _installenabled = (Net40Present
                    && (IIS6Present || IIS7Present)
                    && (Sql2008Present || Sql2012Present || SqlExpressPresent));

                return _installenabled;
            }
            set
            {
                _installenabled = value;
            }
        }
        public bool SqlExpressPresent
        {
            get;
            set;
        }
        public bool SqlExpressAbsent
        {
            get
            {
                return !SqlExpressPresent;
            }
        }

        public bool Sql2008Present
        {
            get;
            set;
        }
        public bool Sql2008Absent
        {
            get
            {
                return !Sql2008Present;
            }
        }
        public bool Sql2012Present
        {
            get;
            set;
        }
        public bool Sql2012Absent
        {
            get { return !Sql2012Present; }
        }

        public bool IIS6Present
        {
            get;
            set;
        }
        public bool IIS6Absent
        {
            get
            {
                return !IIS6Present;
            }
        }
        public bool IIS7Present
        {
            get;
            set;
        }
        public bool IIS7Absent
        {
            get { return !IIS7Present; }
        }

        public bool Net40Present
        {
            get;
            set;
        }
        public bool Net40Absent
        {
            get
            {
                return !Net40Present;
            }
        }
        #endregion
    }
}
