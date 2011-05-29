using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using KinectSlideControl.Windows.Properties;
using Microsoft.Office.Core;
using ppt = Microsoft.Office.Interop.PowerPoint;

namespace KinectSlideControl.Windows
{
    public partial class Form1 : Form
    {
        private PowerPointHelper helper;

        public Form1()
        {
            InitializeComponent();
        }


        private void btnOpen_Click(object sender, EventArgs e)
        {
            openSlide();
        }

        private void openSlide()
        {
            string path = getFilePath();
            
            if (path == null) { return; }

            helper.OpenPowerPointApplication(path);
            helper.SetSlideConfiguration();
            helper.ShowSlidePresentation();
        }

        private string getFilePath()
        {
            DialogResult dialogResult=  openPptDialog.ShowDialog(this);

            if (dialogResult == DialogResult.OK)
            {
                return openPptDialog.FileName;
            }

            return null;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            helper.PrevSlide();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            helper.NextSlide();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            helper.ClosePresentation();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openSlide();
        }

        private void fecharToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            helper = new PowerPointHelper();
            this.Hide();

            notifyIcon1.Icon = Resources.Openkinect_thumbnail;
            notifyIcon1.ShowBalloonTip(1000);
        }

        private void restaurarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
        }
    }
}
