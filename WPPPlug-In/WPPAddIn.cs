using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AddInDesignerObjects;
using Office;
using PowerPoint;
using WPPPlug_In.Properties;

namespace WPPPlug_In
{
    public class WPPAddIn:IDTExtensibility2,IRibbonExtensibility
    {
        public PowerPoint.Application application = null;
        public void OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            this.application = application as PowerPoint.Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if(application != null)
            {
                application = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnStartupComplete(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }

        public string GetCustomUI(string RibbonID)
        {
            return Resources.WPPRibbon;
        }

        public Image GetRibbonImage(string imageName)
        {
            switch (imageName)
            {
                case "Project":
                    return Resources.project;
                case "Calculate":
                    return Resources.calculate;
            }
            return null;
        }

        public void PencilButton_Click(IRibbonControl ribbonControl)
        {
            MessageBox.Show("选中了铅笔");
        }

        public void WaterColorButton_Click(IRibbonControl ribbonControl)
        {
            MessageBox.Show("选中了水彩笔");
        }

        public void OkButton_Click(IRibbonControl ribbonControl)
        {
            MessageBox.Show("点击了确定");
        }
    }
}
