using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace MyOBCustomService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

            //#if false

            //            MyOBCustomService cs = new MyOBCustomService();
            //            cs.OnDebug();
            //#else
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new MyOBCustomService()
            };
            ServiceBase.Run(ServicesToRun);
       }
//#endif
//        }
    }
}
