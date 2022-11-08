# escape=`
FROM mcr.microsoft.com/windows/server:ltsc2022
ENV CHOCO_URL=https://chocolatey.org/install.ps1
RUN powershell -Command "Set-ExecutionPolicy Bypass -Scope Process -Force; `
 [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Tls,Tls11,Tls12'; `
 iex ((New-Object System.Net.WebClient).DownloadString("$env:CHOCO_URL"));"
RUN powershell -Command "choco install -y office-to-pdf; choco install -y office-tool"

WORKDIR C:\\odtsetup
COPY ./odtsetup/setup.exe .
COPY ./Office ./Office
ADD config.xml .
RUN setup.exe /configure C:\\odtsetup\\config.xml

WORKDIR /
COPY loginoffice.ps1 C:\\loginoffice.ps1
COPY entrypoint.cmd C:\\entrypoint.cmd
COPY test.ps1 C:\\test.ps1
ENV OFFICEUSER="null"
ENV OFFICEPASSWD="null"
RUN rmdir /s /q C:\\odtsetup
RUN powershell -Command "new-object -comobject word.application"
RUN mkdir C:\Windows\System32\config\systemprofile\Desktop
RUN powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; `
  Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force"
SHELL [ "powershell", "-Command"]
RUN Install-Module AzureAD -Confirm:$false -Force; Install-Module MSOnline -Confirm:$false -Force
ENTRYPOINT [ "C:\\entrypoint.cmd" ]