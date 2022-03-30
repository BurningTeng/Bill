using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenCvSharp;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using Titanium.Web.Proxy;
using Titanium.Web.Proxy.EventArguments;
using Titanium.Web.Proxy.Models;

namespace WorkerService1
{
    internal class SignImplement
    {

        private readonly ProxyServer proxyServer = new();
        public bool ExitFlag { get; set; }

        public string target = "";
        public string template = "";
        public static int Match(string template, string target)
        {
            //模板图片
            Mat temp = new Mat(template, ImreadModes.AnyColor);
            //被匹配图
            Mat wafer = new Mat(target, ImreadModes.AnyColor);

            //Canny边缘检测
            Mat temp_canny_Image = new();
            Cv2.Canny(temp, temp_canny_Image, 100, 200);
            Mat wafer_canny_Image = new();
            Cv2.Canny(wafer, wafer_canny_Image, 100, 200);

            //匹配结果
            Mat result = new();
            //模板匹配
            Cv2.MatchTemplate(wafer_canny_Image, temp_canny_Image, result, TemplateMatchModes.CCoeffNormed);//最好匹配为1,值越小匹配越差
            Cv2.MinMaxLoc(result, out _, out Point maxLoc);
            Point matchLoc = maxLoc;

            return matchLoc.X + 16;
        }

        public static void SendMail(string content, string filepath, string email)
        {
            MailMessage message = new();

            //设置发件人,发件人需要与设置的邮件发送服务器的邮箱一致
            MailAddress fromAddr = new("2250911301@qq.com");
            message.From = fromAddr;
            //设置收件人,可添加多个,添加方法与下面的一样
            message.To.Add(email);
            //设置抄送人
            // message.CC.Add("1592035782@qq.com");
            //设置邮件标题
            message.Subject = "打卡通知";
            //设置邮件内容
            message.Body = content;
            //添加附件
            Attachment data = new(filepath, MediaTypeNames.Application.Octet);
            // Add time stamp information for the file.
            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(filepath);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(filepath);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(filepath);
            // Add the file attachment to this email message.
            message.Attachments.Add(data);
            //设置邮件发送服务器,服务器根据你使用的邮箱而不同,可以到相应的邮箱管理后台查看
            SmtpClient client = new SmtpClient("smtp.qq.com", 25);
            //设置发送人的邮箱账号和授权码
            client.Credentials = new NetworkCredential("2250911301@qq.com", "gideydwxiwmudjji");

            //启用ssl,也就是安全发送
            client.EnableSsl = true;
            //发送邮件
            client.Send(message);
        }

        public void Sign(string name, string email, string password)
        {
            IWebDriver wd = new ChromeDriver();
            wd.Navigate().GoToUrl("http://kq.neusoft.com/");
            IWindow window = wd.Manage().Window;
            window.Maximize();
            Thread.Sleep(3000);
            wd.FindElement(By.ClassName("userName")).SendKeys(name);
            wd.FindElement(By.ClassName("password")).SendKeys(password);
            int distance = 0;
            while (true)
            {
                Monitor.Enter(this);
                distance = Match(AppDomain.CurrentDomain.BaseDirectory + "\\template.png", AppDomain.CurrentDomain.BaseDirectory + "\\target.png");
                Monitor.Exit(this);
                //这是滑块
                var slide = wd.FindElement(By.ClassName("ui-slider-btn"));
                Actions action = new Actions(wd);
                //点击并按住滑块元素
                action.ClickAndHold(slide);
                action.MoveByOffset(distance, 0);
                string alert;

                try
                {
                    action.Release().Perform();
                    Thread.Sleep(2000);
                    alert = wd.FindElement(By.ClassName("ui-slider-text")).Text;
                }
                catch (Exception e)
                {
                    Console.WriteLine("发生异常:" + e.ToString());
                    alert = "";
                }

                if (alert.Contains("验证成功"))
                {
                    break;
                }
                else
                {
                    wd.SwitchTo().DefaultContent();
                }
                Thread.Sleep(2000);
            }

            wd.FindElement(By.Id("loginButton")).Click();
            Thread.Sleep(3000);

            string js_sign = "javascript:document.attendanceForm.submit();";
            ((ChromeDriver)wd).ExecuteScript(js_sign, null);
            Thread.Sleep(2000);
            //截屏
            var screenshot = ((ChromeDriver)wd).GetScreenshot();
            System.Drawing.Image screenshotImage;
            using (MemoryStream memStream = new MemoryStream(screenshot.AsByteArray))
            {
                screenshotImage = System.Drawing.Image.FromStream(memStream);
            }
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "screenshot.png";
            screenshotImage.Save(filePath);

            //发送邮件
            SendMail("打卡成功" + DateTime.Now, filePath, email);
            //调用js
            string js_exit = "javascript:exitAttendance();";
            //设置ExitFlag，避免在退出的时候重新下载图片
            ExitFlag = true;
            ((ChromeDriver)wd).ExecuteScript(js_exit, null);
            Thread.Sleep(3000);
            wd.Quit();
        }

        // locally trust root certificate used by this proxy 
        //proxyServer.CertificateManager.TrustRootCertificate(true);

        // optionally set the Certificate Engine
        // Under Mono only BouncyCastle will be supported
        //proxyServer.CertificateManager.CertificateEngine = Network.CertificateEngine.BouncyCastle;
        public void StartProxyServer()
        {
            proxyServer.BeforeRequest += OnRequest;
            proxyServer.BeforeResponse += OnResponse;
            proxyServer.ServerCertificateValidationCallback += OnCertificateValidation;
            proxyServer.ClientCertificateSelectionCallback += OnCertificateSelection;
        }

        public void SetProxyPort(int port)
        {
            var explicitEndPoint = new ExplicitProxyEndPoint(IPAddress.Any, port, true)
            {
                // Use self-issued generic certificate on all https requests
                // Optimizes performance by not creating a certificate for each https-enabled domain
                // Useful when certificate trust is not required by proxy clients
                //GenericCertificate = new X509Certificate2(Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "genericcert.pfx"), "password")
            };

            // Fired when a CONNECT request is received
            //explicitEndPoint.BeforeTunnelConnect += OnBeforeTunnelConnect;

            // An explicit endpoint is where the client knows about the existence of a proxy
            // So client sends request in a proxy friendly manner
            proxyServer.AddEndPoint(explicitEndPoint);
            proxyServer.Start();

            //proxyServer.UpStreamHttpProxy = new ExternalProxy() { HostName = "localhost", Port = 8888 };
            //proxyServer.UpStreamHttpsProxy = new ExternalProxy() { HostName = "localhost", Port = 8888 };
            foreach (var endPoint in proxyServer.ProxyEndPoints)
                Console.WriteLine("Listening on '{0}' endpoint at Ip {1} and port: {2} ",
                    endPoint.GetType().Name, endPoint.IpAddress, endPoint.Port);

            // Only explicit proxies can be set as system proxy!
            proxyServer.SetAsSystemHttpProxy(explicitEndPoint);
            proxyServer.SetAsSystemHttpsProxy(explicitEndPoint);
        }

        public void StopProxyServer()
        {
            // Unsubscribe & Quit
            //explicitEndPoint.BeforeTunnelConnect -= OnBeforeTunnelConnect;
            proxyServer.BeforeRequest -= OnRequest;
            proxyServer.BeforeResponse -= OnResponse;
            proxyServer.ServerCertificateValidationCallback -= OnCertificateValidation;
            proxyServer.ClientCertificateSelectionCallback -= OnCertificateSelection;
            proxyServer.Stop();
        }

        private async Task OnRequest(object sender, SessionEventArgs e)
        {
            var method = e.HttpClient.Request.Method.ToUpper();
            if ((method == "POST" || method == "PUT" || method == "PATCH"))
            {
                // Get/Set request body bytes
                byte[] bodyBytes = await e.GetRequestBody();
                e.SetRequestBody(bodyBytes);

                // Get/Set request body as string
                string bodyString = await e.GetRequestBodyAsString();
                e.SetRequestBodyString(bodyString);

                // store request 
                // so that you can find it from response handler 
                e.UserData = e.HttpClient.Request;
            }

            // To cancel a request with a custom HTML content
            // Filter URL
            if (e.HttpClient.Request.RequestUri.AbsoluteUri.Contains("360"))
            {
                e.Ok("<!DOCTYPE html>" +
                    "<html><body><h1>" +
                    "Website Blocked" +
                    "</h1>" +
                    "<p>Blocked by titanium web proxy.</p>" +
                    "</body>" +
                    "</html>");
            }

            // Redirect example
            if (e.HttpClient.Request.RequestUri.AbsoluteUri.Contains("wikipedia.org"))
            {
                e.Redirect("https://www.paypal.com");
            }
        }

        // Modify response
        private async Task OnResponse(object sender, SessionEventArgs e)
        {
            //if (!e.ProxySession.Request.Host.Equals("medeczane.sgk.gov.tr")) return;
            if (e.HttpClient.Request.Method == "GET" || e.HttpClient.Request.Method == "POST")
            {
                if (e.HttpClient.Response.StatusCode == 200)
                {
                    string stringResponse = await e.GetResponseBodyAsString();
                    if ("http://kq.neusoft.com/jigsaw".Equals(e.HttpClient.Request.Url))
                    {
                        //exit的时候会走第二次。 在exit之前调用StopProxyServer，防止出现第二次走这个方法的情况。
                        if (!ExitFlag)
                            SaveImage(stringResponse);
                    }
                }
            }
        }

        private void SaveImage(String stringResponse)
        {
            var jo = JsonConvert.DeserializeObject(stringResponse) as JObject;
            template = jo["smallImage"].ToString();
            target = jo["bigImage"].ToString();
            template = "http://kq.neusoft.com/upload/jigsawImg/" + template + ".png";
            target = "http://kq.neusoft.com/upload/jigsawImg/" + target + ".png";

            WebClient client = new();
            //不加锁的话只能下载第一个图片，然后就去匹配去了，由于第二个图片还没有下载下来，导致匹配的时候报错。
            //为什么第二个图片下载不下来需要进一步调查。
            Monitor.Enter(this);
            client.DownloadFile(target, AppDomain.CurrentDomain.BaseDirectory + "\\target.png");
            client.DownloadFile(template, AppDomain.CurrentDomain.BaseDirectory + "\\template.png");
            Monitor.Exit(this);
        }

        // Allows overriding default certificate validation logic
        private Task OnCertificateValidation(object sender, CertificateValidationEventArgs e)
        {
            // set IsValid to true/false based on Certificate Errors
            if (e.SslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
                e.IsValid = true;

            return Task.CompletedTask;
        }

        // Allows overriding default client certificate selection logic during mutual authentication
        private Task OnCertificateSelection(object sender, CertificateSelectionEventArgs e)
        {
            // set e.clientCertificate to override
            return Task.CompletedTask;
        }
    }
}