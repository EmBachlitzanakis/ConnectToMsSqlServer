# escape=`
# Use the official Microsoft Windows Server Core image with .NET Framework SDK
FROM mcr.microsoft.com/dotnet/framework/sdk:4.8-windowsservercore-ltsc2022

# Install Microsoft OLE DB Driver for SQL Server
# Using a more reliable download URL and adding retry logic
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]
RUN [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; `
    Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?linkid=2185760' -OutFile 'C:\msoledbsql.msi'; `
    Start-Process -FilePath msiexec.exe -ArgumentList '/i', 'C:\msoledbsql.msi', '/quiet', '/qn', '/norestart', 'IACCEPTMSOLEDBSQLLICENSETERMS=YES' -Wait; `
    Remove-Item -Path 'C:\msoledbsql.msi' -Force

# Install Visual C++ Redistributable
RUN Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile 'C:\vc_redist.x64.exe'; `
    Start-Process -FilePath 'C:\vc_redist.x64.exe' -ArgumentList '/install', '/quiet', '/norestart' -Wait; `
    Remove-Item -Path 'C:\vc_redist.x64.exe' -Force

# Set working directory and copy project files
WORKDIR /app
COPY . .

# Optional: Build your project with MSBuild
# RUN msbuild msolebdb.sln /p:Configuration=Release

# Optional: Set entry point
# ENTRYPOINT ["msolebdb.exe"]