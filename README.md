###  通过BackgroundService来创建Windows服务
核心代码是自动打卡系统功能，使用的实现方式是window服务。
在https://github.com/BurningTeng/NeusoftSign有控制台版本的实现。
欢迎一起讨论交流。
###  安装和启动windows服务
sc.exe create "SmartSign2" binpath="D:\vs2022\Project\WorkerService1\bin\Release\net6.0\WorkerService1.exe"
sc.exe failure "SmartSign2" reset=0 actions=restart/60000/restart/60000/run/1000
sc qfailure "SmartSign2"
sc.exe start "SmartSign2"