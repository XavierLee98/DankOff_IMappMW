using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using IMAppSapMidware_NetCore.Helper;
using System.Threading;
using System.Threading.Tasks;

namespace IMAppSapMidware_NetCore
{
    public class Worker : BackgroundService
    {
        //readonly ILogger<Worker> _logger;
        readonly IConfiguration _configuration;
        readonly string _dbConnStrSectionKey_Midware = "Midware";
        readonly string _dbConnStrSectionKey_Erp = "Erp";

        string _DbConnectStrMidware { get; set; } = string.Empty;
        string _DbConnectStrErp { get; set; } = string.Empty;
        string _ErpDbName { get; set; } = string.Empty;
        static int _timerInterval { get; set; } = 3000;
        static System.Timers.Timer _checkTimer { get; set; } = new System.Timers.Timer(_timerInterval); // three second

        public CancellationToken cancellationToken { get; set; } = new CancellationToken();

        public Worker(ILogger<Worker> logger, IConfiguration configuration)
        {
            //_logger = logger;
            _configuration = configuration;
            _DbConnectStrMidware = configuration.GetConnectionString(_dbConnStrSectionKey_Midware);
            _DbConnectStrErp = configuration.GetConnectionString(_dbConnStrSectionKey_Erp);
            _ErpDbName = configuration.GetSection("AppSettings").GetSection("ErpDb").Value;

            Program._DbErpConnStr = _DbConnectStrErp;
            Program._DbMidwareConnStr = _DbConnectStrMidware;
            Program._ErpDbName = _ErpDbName;

            InitialTimer();
            StartTimer();
        }

        void InitialTimer()
        {
            if (_checkTimer == null) _checkTimer = new System.Timers.Timer(_timerInterval);
            _checkTimer.Elapsed += _checkTimer_Elapsed;
        }

        void _checkTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            StopTimer();
            //Program.FilLogger?.Log("Check and execute request(s).");
            var reqHelper = new RequestsHelper
            {
                DbConnectString_Erp = this._DbConnectStrErp,
                DbConnectString_Midware = this._DbConnectStrMidware,
                LasteErrorMessage = string.Empty
            };

            reqHelper.ExecuteRequest();

            Thread.Sleep(600);
            StartTimer();
        }

        void StartTimer()
        {
            _checkTimer.Enabled = true;
            _checkTimer.Start();
            //_logger.LogInformation("Timer started at: {time}", DateTimeOffset.Now);
            // Program.FilLogger?.Log("timer started.");
        }

        void StopTimer()
        {
            _checkTimer.Enabled = false;
            _checkTimer.Stop();
            // Program.FilLogger?.Log("timer stopped.");
            //_logger.LogInformation("Timer stop at: {time}", DateTimeOffset.Now);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // to do code
            // no used use to sycnchorone process
        }
    }
}
