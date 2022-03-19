namespace WorkerService1;

public sealed class WindowsBackgroundService : BackgroundService
{
    private readonly ILogger<WindowsBackgroundService> _logger;

    public WindowsBackgroundService(ILogger<WindowsBackgroundService> logger)
    {
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            SignFilter();
            await Task.Delay(1000 * 60, stoppingToken);
        }
    }

    private void SignFilter()
    {
        Random random = new Random();
        int min = 0;
        int max = 120;
        SignImplement si = new SignImplement();

        if (DateTimeOffset.Now.Hour == 7 && DateTimeOffset.Now.Minute == 00)  //如果当前时间是20点15分
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            si.StartProxyServer();
            si.SetProxyPort();
            var task = Task.Run(() =>
            {
                si.Sign();
            });
            task.Wait();
            _logger.LogWarning("早上打卡结束");
        }
        else if (DateTimeOffset.Now.Hour == 19 && DateTimeOffset.Now.Minute == 00)
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            si.StartProxyServer();
            si.SetProxyPort();
            var task = Task.Run(() =>
            {
                si.Sign();
            });
            task.Wait();
            _logger.LogWarning("晚上打卡结束");
        } 
        else
        {
            _logger.LogWarning("当前时间是: {time}", DateTimeOffset.Now);
        }
    }
}