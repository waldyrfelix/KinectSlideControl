
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace KinectSlideControl.Windows
{
    public class PowerPointHelper
    {
        private Application application;
        private Presentation presentation;
        private SlideShowWindow window;

        public void OpenPowerPointApplication(string path)
        {
            application = new Application();
            presentation = application.Presentations.Open(path, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
        }

        public void ShowSlidePresentation()
        {
            window = presentation.SlideShowSettings.Run();
        }

        public void SetSlideConfiguration()
        {
            presentation.SlideShowSettings.StartingSlide = 1;
            presentation.SlideShowSettings.EndingSlide = presentation.Slides.Count;
        }

        public void PrevSlide()
        {
            if (window != null)
            {
                window.View.Previous(); 
            }
        }

        public void NextSlide()
        {
            if (window != null)
            {
                window.View.Next();
            }
        }

        public void ClosePresentation()
        {
            if (presentation != null)
            {
                presentation.Close(); 
            }
        }
    }
}
