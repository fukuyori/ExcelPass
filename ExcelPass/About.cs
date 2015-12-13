using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment.Application;

namespace WindowsFormsApplication1 {
    public partial class About : Form {
        public About() {
            InitializeComponent();

            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            label2.Text = "Ver." + Application.ProductVersion;
        }

        private void button1_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
