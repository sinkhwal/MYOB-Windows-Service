using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace MyOBCustomService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            
            InitializeComponent();
         
            //this.AfterInstall += new InstallEventHandler(serviceInstaller1_AfterInstall);
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
        
            //using (ServiceController sc = new ServiceController(serviceInstaller1.ServiceName))
            //{
            //    sc.Start();
            //}
        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}
