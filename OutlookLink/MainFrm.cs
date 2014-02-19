using System;
using System.Drawing;
using System.Windows.Forms;
using OutlookLink.Properties;

namespace OutlookLink
{
    public partial class MainFrm : Form
    {
        public MainFrm()
        {
            InitializeComponent();

            this.Icon = Resources.outlook;
            this.notifyIcon1.Icon = Resources.outlook;
        }

        private void registerToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            RegisterTool.Register("outlook", System.Windows.Forms.Application.ExecutablePath, "");
        }

        private void notifyIcon1_Click(object sender, System.EventArgs e)
        {
            Point position = Cursor.Position;
            Point menuPoint = new Point();
            menuPoint.X = position.X - this.outlookLinkMenuStrip.Size.Width;
            menuPoint.Y = position.Y - this.outlookLinkMenuStrip.Size.Height;
            this.outlookLinkMenuStrip.Show(menuPoint);
        }

        private void MainFrm_Load(object sender, EventArgs e)
        {
            RegisterTool.Register("outlook", System.Windows.Forms.Application.ExecutablePath, "");
            this.Close();
        }
    }
}
