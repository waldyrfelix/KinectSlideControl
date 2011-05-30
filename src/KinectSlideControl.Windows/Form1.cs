using System;
using System.Windows.Forms;
using KinectSlideControl.Windows.Properties;

namespace KinectSlideControl.Windows
{
    public partial class Form1 : Form
    {
        private PowerPointHelper pptHelper;
        private KinectHelper kinectHelper;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Hide();

            notifyIcon1.Icon = Resources.Openkinect_thumbnail;
            notifyIcon1.ShowBalloonTip(500);

            pptHelper = new PowerPointHelper();
            kinectHelper = new KinectHelper();

            kinectHelper.NextSlide += new NextSlideEventHandler(kinectHelper_NextSlide);
            kinectHelper.PrevSlide += new PrevSlideEventHandler(kinectHelper_PrevSlide);
            kinectHelper.IniciarKinect();

        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            openSlide();
        }

        private void openSlide()
        {
            string path = getFilePath();
            
            if (path == null) { return; }

            pptHelper.OpenPowerPointApplication(path);
            pptHelper.SetSlideConfiguration();
            pptHelper.ShowSlidePresentation();
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
            pptHelper.PrevSlide();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            pptHelper.NextSlide();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            pptHelper.ClosePresentation();
            kinectHelper.Close();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openSlide();
        }

        private void fecharToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void kinectHelper_PrevSlide()
        {
            pptHelper.PrevSlide();
        }

        void kinectHelper_NextSlide()
        {
            pptHelper.NextSlide();
        }

        private void restaurarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
        }
    }
}
