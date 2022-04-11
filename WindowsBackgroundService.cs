namespace WorkerService1;

using ThirdParty.Json.LitJson;

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
        Random random = new();
        int min = 0;
        int max = 120;
        SignImplement si = new();

        var listPerson = JsonMapper.ToObject<List<PersonInfo>>(File.ReadAllText("PersonInfo.json"));


        if (DateTimeOffset.Now.Hour == 7 && DateTimeOffset.Now.Minute == 00)  //如果当前时间是7点00分
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            si.Sign(listPerson[0].name, listPerson[0].email, listPerson[0].password);
            _logger.LogWarning("早上打卡结束");
        }
        else if (DateTimeOffset.Now.Hour == 19 && DateTimeOffset.Now.Minute == 00)
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            si.Sign(listPerson[0].name, listPerson[0].email, listPerson[0].password);
            _logger.LogWarning("晚上打卡结束");
        } 
        else
        {
            _logger.LogWarning("当前时间是: {time}", DateTimeOffset.Now);
        }
    }
}