using System;
using System.ServiceProcess;
using System.Threading;

namespace BotGestaoDefeitos
{
    public partial class Service1 : ServiceBase
    {
        private ManualResetEvent _shutdownEvent = new ManualResetEvent(false);
        private Thread _thread;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
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
                    
                    log4net.LogManager.GetLogger("Processamento.Geral.Info").Info("Iniciando gestão de defeitos");
                    new GestaoDefeitos().ExecutarGestaoDefeitos();
                }
                catch (Exception ex)
                {
                    log4net.LogManager.GetLogger("Processamento.Geral.Erro").Error($"Falha ao realizar gestão de defeitos.", ex);
                }
                finally
                {
                    System.Threading.Thread.Sleep(6000);
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
