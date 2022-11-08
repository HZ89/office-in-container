# How to make office image

## 概述

1. 安装
   dockerfile 只能通过运行 cmd 或者 powershell 来构建容器镜像，这里只能选择非交互式的安装方法。windows官方提供了 [office deploy tools](https://learn.microsoft.com/en-us/deployoffice/overview-office-deployment-tool). 通过一个xml配置可以静默下载安装office。

   经过测试 windows 2019 版本的容器并不能成功执行此部署工具。虽然手工启动临时layer之后执行不报错，但是安装完毕之后无法调用到word的 com object

   windows 2022 可以顺利安装。之后可以调用到 com object (powershell 执行 `$word = New-Object -ComObject word.application` 来new 一个word)

2. 激活
   常规 office 分为 零售版，批量授权 kms版本，批量key授权版本。一般用kms+voulme版本来激活。微软同样提供了 [ospp.vbc](https://learn.microsoft.com/en-us/deployoffice/vlactivation/tools-to-manage-volume-activation-of-office) 来在命令行执行激活操作。
   ospp脚本在2022 windows上执行可以查询激活状态，设置kms server地址，但是执行 激活命令会报错。并且error code 从微软的error code目录里面查询是一个windows激活的错误，与office无关。猜测是激活office之前先校验了os的激活状态。容器内可能缺少了外部windows的激活信息。

3. office 365
   office 365 为订阅制服务，无需进行激活，但是需要登录微软的账号，windows 提供了命令行登录并管理的工具 [AzureAD](https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell?view=o365-worldwide).以及 [MSOnline](https://learn.microsoft.com/en-us/powershell/module/msonline/?view=azureadps-1.0)

## How to use this project

1. 下载[office deploy tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117), 命令中的版本可能过期，请从超链下载最新版

   ```powershell
   Invoke-WebRequest "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_15629-20208.exe" -OutFile ./odt.exe
   ./odt.exe /quiet /norestart /extract:./odtsetup
   ```

2. 当前目录下载office:
  
    ```powershell
    ./odtsetup/setup.exe /download ./config.xml
    ls ./ # it is will download office 365 into local Office dir
    ```

3. build:

    ```powershell
    docker build -t test:v1 .
    ```

4. run:

    ```powershell
    docker run -ti -e "OFFICEUSER=XXXXX" -e "OFFICEPASSWD=XXXXXX" test:v1 powershell
    ```
