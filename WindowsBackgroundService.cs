namespace WorkerService1;

using System.Collections;
using System.Net;
using System.Net.NetworkInformation;
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

    public int GetFreePort()
    {
        int port = 0;
        while (true)
        {
            int start = 1024;
            int end = 65535;
            var random = new Random();
            port = random.Next(start, end);
            if (!PortIsUsed().Contains(port))
            {
                Console.WriteLine(port);
                break;
            }
        }
        return port;
    }

    /// <summary>        
    /// 获取操作系统已用的端口号        
    /// </summary>        
    /// <returns></returns>
    /// https://www.cnblogs.com/xdoudou/p/3605134.html#:~:text=C%23%E9%9A%8F%E6%9C%BA%E5%8F%96%E5%BE%97%E5%8F%AF%E7%94%A8%E7%AB%AF%E5%8F%A3%E5%8F%B7%20TCP%E4%B8%8EUDP%E6%AE%B5%E7%BB%93%E6%9E%84%E4%B8%AD%E7%AB%AF%E5%8F%A3%E5%9C%B0%E5%9D%80%E9%83%BD%E6%98%AF16%E6%AF%94%E7%89%B9%EF%BC%8C%E5%8F%AF%E4%BB%A5%E6%9C%89%E5%9C%A80---65535%E8%8C%83%E5%9B%B4%E5%86%85%E7%9A%84%E7%AB%AF%E5%8F%A3%E5%8F%B7%E3%80%82,%E5%AF%B9%E4%BA%8E%E8%BF%9965536%E4%B8%AA%E7%AB%AF%E5%8F%A3%E5%8F%B7%E6%9C%89%E4%BB%A5%E4%B8%8B%E7%9A%84%E4%BD%BF%E7%94%A8%E8%A7%84%E5%AE%9A%EF%BC%9A%20%EF%BC%881%EF%BC%89%E7%AB%AF%E5%8F%A3%E5%8F%B7%E5%B0%8F%E4%BA%8E256%E7%9A%84%E5%AE%9A%E4%B9%89%E4%B8%BA%E5%B8%B8%E7%94%A8%E7%AB%AF%E5%8F%A3%EF%BC%8C%E6%9C%8D%E5%8A%A1%E5%99%A8%E4%B8%80%E8%88%AC%E9%83%BD%E6%98%AF%E9%80%9A%E8%BF%87%E5%B8%B8%E7%94%A8%E7%AB%AF%E5%8F%A3%E5%8F%B7%E6%9D%A5%E8%AF%86%E5%88%AB%E7%9A%84%E3%80%82
    public static IList PortIsUsed()
    {
        //获取本地计算机的网络连接和通信统计数据的信息            
        IPGlobalProperties ipGlobalProperties = IPGlobalProperties.GetIPGlobalProperties();
        //返回本地计算机上的所有Tcp监听程序            
        IPEndPoint[] ipsTCP = ipGlobalProperties.GetActiveTcpListeners();
        //返回本地计算机上的所有UDP监听程序            
        IPEndPoint[] ipsUDP = ipGlobalProperties.GetActiveUdpListeners();
        //返回本地计算机上的ipv4/ipv6传输控制协议(TCP)连接的信息。            
        TcpConnectionInformation[] tcpConnInfoArray = ipGlobalProperties.GetActiveTcpConnections();

        IList allPorts = new ArrayList();
        foreach (IPEndPoint ep in ipsTCP)
        {
           allPorts.Add(ep.Port);
        }
        foreach (IPEndPoint ep in ipsUDP)
        {
            allPorts.Add(ep.Port);
        }
        foreach (TcpConnectionInformation conn in tcpConnInfoArray)
        {
            allPorts.Add(conn.LocalEndPoint.Port);
        }
        return allPorts;
    }

    private void SignFilter()
    {
        Random random = new();
        int min = 0;
        int max = 120;
        SignImplement si = new();

        var listPerson = JsonMapper.ToObject<List<PersonInfo>>(File.ReadAllText("PersonInfo.json"));


        if (DateTimeOffset.Now.Hour == 7 && DateTimeOffset.Now.Minute == 00)  //如果当前时间是20点15分
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            var task = Task.Run(() =>
            {
                si.StartProxyServer();
                si.SetProxyPort(GetFreePort());
                si.Sign(listPerson[0].name, listPerson[0].email, listPerson[0].password);
            });
            task.Wait();
            _logger.LogWarning("早上打卡结束");
        }
        else if (DateTimeOffset.Now.Hour == 19 && DateTimeOffset.Now.Minute == 00)
        {
            Thread.Sleep(random.Next(min, max) * 60 * 1000);
            var task = Task.Run(() =>
            {
                si.StartProxyServer();
                si.SetProxyPort(GetFreePort());
                si.Sign(listPerson[0].name, listPerson[0].email, listPerson[0].password);
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