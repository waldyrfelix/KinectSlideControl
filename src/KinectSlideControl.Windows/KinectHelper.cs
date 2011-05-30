using System.Diagnostics;
using System.Threading;
using ManagedNite;

namespace KinectSlideControl.Windows
{
    public delegate void NextSlideEventHandler();

    public delegate void PrevSlideEventHandler();

    public class KinectHelper
    {
        private XnMOpenNIContext context;
        private bool neverStarted;
        private Thread thread;
        private XnMSessionManager sessionManager;
        private XnMSwipeDetector swipeDetector;

        public event NextSlideEventHandler NextSlide;
        public event PrevSlideEventHandler PrevSlide;

        public void IniciarKinect()
        {
            Debug.WriteLine("Iniciando acceso a Kinect ...");

            //Instancia o contexto do Kinect para Iniciar Operação
            context = new XnMOpenNIContext();

            //Inicia Kinect
            context.Init();
            neverStarted = true;

            //SessionManager configurado para quando você sacudir a mão ele iniciar 
            //Wave seria o movimento de tchau
            //E RaiseHand é só erguer a mão...
            //Os eventos serão disparador com o foco na mão, ou seja, se perder o foco vai ter que fazer o RaiseHand novamente
            sessionManager = new XnMSessionManager(context, "Wave", "RaiseHand");


            swipeDetector = new XnMSwipeDetector(true);
            swipeDetector.MotionTime = 600;
            swipeDetector.SteadyDuration = 10; //200   uint
            swipeDetector.SteadyMaxVelocity = 1f; //0.01    float
            swipeDetector.XAngleThreshold = 25.00f; //25.0      float
            swipeDetector.YAngleThreshold = 25.00f; //20.0      float  
            swipeDetector.GeneralSwipe += OnGeneralSwipe;

            sessionManager.AddListener(swipeDetector);

            //O evento é disparado quando for feito o tchauzinho
            sessionManager.SessionStarted += SessionManagerSessionStarted;

            //Thread Necessária
            thread = new Thread(mainLoop);
            thread.Start();

            //Mensagem de Inicio
            Debug.WriteLine("Kinect configurado. Fazer saudação para iniciar Sessao");
            Debug.WriteLine("");
        }

        protected void SessionManagerSessionStarted(object sender, PointEventArgs e)
        {
            if (neverStarted)
            {
                Debug.WriteLine("Sessão Startada");
                neverStarted = false;
            }
            else
            {
                Debug.WriteLine("Sessão Retomada");
            }
        }


        protected void OnGeneralSwipe(object sender, SwipeDetectorGeneralEventArgs e)
        {
            if (e.SelectDirection == Direction.Left)
            {
                Debug.WriteLine("Left");
                if (PrevSlide != null)
                {
                    PrevSlide();
                }
            }
            else if (e.SelectDirection == Direction.Right)
            {
                Debug.WriteLine("Right");
                if (NextSlide != null)
                {
                    NextSlide();
                }
            }
        }

        public void Close()
        {
            try
            {
                this.thread.Abort();
            }
            catch (ThreadAbortException) { }
        }

        private void mainLoop()
        {
            while (true)
            {
                context.Update();
                sessionManager.Update(context);
            }
        }
    }
}