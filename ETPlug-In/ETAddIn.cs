using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AddInDesignerObjects;
using Office;
using Excel;

namespace ETPlug_In
{
    public class ETAddIn: IDTExtensibility2, IRibbonExtensibility
    {
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnStartupComplete(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public string GetCustomUI(string RibbonID)
        {
            //throw new NotImplementedException();
            return "";
        }
    }
}
