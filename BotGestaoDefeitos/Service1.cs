using log4net;
using System;
using System.ServiceProcess;
using System.Threading;

namespace BotGestaoDefeitos
{
    public partial class Service1 : ServiceBase
    {
        private ManualResetEvent _shutdownEvent = new ManualResetEvent(false);
        private Thread _thread;

        private static readonly ILog logInfo = LogManager.GetLogger(typeof(Program));
        private static readonly ILog logErro = LogManager.GetLogger(typeof(Program));
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            _thread = new Thread(StartBackGround);

            _thread.Name = "Background";
            _thread.IsBackground = true;
            _thread.Start();
        }
        public void Start()
        {
            StartBackGround(null);
        }

        private void StartBackGround(object obj)
        {
            //ThreadPool.QueueUserWorkItem(a => GestaoDefeitos());
            Thread threadGestaoDefeitos = new Thread(GestaoDefeitos);
            threadGestaoDefeitos.Start();
        }

        protected void GestaoDefeitos()
        {
            while (true)
            {
                try
                {
                    //if (DateTime.Now.Hour == 0 && DateTime.Now.Minute <= 10)
                    //{

                    logInfo.Info("Iniciando gestão de defeitos");
                        new GestaoDefeitos().ExecutarGestaoDefeitos();
                    logInfo.Info("Finalizando gestão de defeitos");
                    //}
                }
                catch (Exception ex)
                {
                    logErro.Error($"Falha ao realizar gestão de defeitos.", ex);
                }
            }
        }

        protected override void OnStop()
        {
            _shutdownEvent.Set();
            if (!_thread.Join(3000))
            {
                _thread.Abort();
            }
        }
    }
}
