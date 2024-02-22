$DaData = (Get-Content "C:\Temp\DadataTokens.json" -Encoding "UTF8") | ConvertFrom-Json
$token = $DaData.token
$secret = $DaData.secret

$flag = (Get-Content "C:\Temp\DataExport.json" -Encoding "UTF8") | ConvertFrom-Json
if ($flag.DataExportSuccess -eq "True")
{
    Write-Host "C:\Temp\DataExport.json is readed DataExportSuccess = True."

    $KKMData = (Get-Content "C:\ProgramData\KKMData.json" -Encoding "UTF8") | ConvertFrom-Json

    Write-Host "C:\ProgramData\KKMData.json is readed."

    if (-not (Test-Path -Path "C:\Temp\DadataInfo.json")) 
    {
        Write-Host "C:\Temp\DadataInfo.json not found."
        $addr = $KKMData[0].pos_address
        $headers = @{
            Accept = 'application/json; charset=utf-8'
            Authorization = 'Token ' + $token
            'X-Secret' = $secret
        }
        $body = ConvertTo-Json @("$addr")
        Write-Host "Sending request to dadata cleaner."
        $response = Invoke-RestMethod 'https://cleaner.dadata.ru/api/v1/clean/address' -Method POST -ContentType "application/json; charset=Windows-1251" -Headers $headers -Body $body
        Write-Host "Request complited."
        ConvertTo-Json -Compress -InputObject $response | Set-Content "C:\Temp\DadataInfo.json" -Encoding "UTF8"
        Write-Host "Saving response to C:\Temp\DadataInfo.json ."
    }
    else
    {
        Write-Host "C:\Temp\DadataInfo.json found. Reading."
        $response = Get-Content "C:\Temp\DadataInfo.json" -Encoding "UTF8" | ConvertFrom-Json
        Write-Host "C:\Temp\DadataInfo.json readed."
        $response = $response.value[0]    
    }

    $iso_code = $response.region_iso_code
    $g_lat = $response.geo_lat
    $g_lon = $response.geo_lon
    Write-Host "Saving POS geo data."
    @{geo_lat="$g_lat"; geo_lon="$g_lon"} | ConvertTo-Json | Set-Content "C:\ProgramData\POS_GEO_DATA.json" -Encoding "UTF8"
    $salt = -join ((65..90) | Get-Random -Count 9 | ForEach-Object {[char]$_})
    $inn = $KKMData[0].inn
    $inn = $inn.Trim()
    $mystream = [IO.MemoryStream]::new([byte[]][char[]]$iso_code)
    $hostmetadata = Get-FileHash -InputStream $mystream -Algorithm MD5
    $hostmetadata = $hostmetadata.hash.tolower()
    
    Set-Location -Path "C:\Temp"
    if ([Environment]::Is64BitOperatingSystem)
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-amd64-openssl.msi"
        Write-Host "its 64-bit OS. Selected $zabbixAgentName."
    }
    else
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-i386-openssl.msi"
        Write-Host "its 32-bit OS. Selected $zabbixAgentName."
    }

    if (Test-Path -Path "C:\Temp\HostName.json")
    {
        Write-Host "C:\Temp\HostName.json found. Reading"
        $hostname = (Get-Content "C:\Temp\HostName.json" -Encoding "UTF8") | ConvertFrom-Json
        $hostname = $hostname.ZabbixHostName
        Write-Host "C:\Temp\HostName.json reading complete. Value - $hostname"
    }
    else
    {
        @{ZabbixHostName="$inn-$iso_code-$salt-POS"} | ConvertTo-Json | Set-Content "C:\Temp\HostName.json" -Encoding "UTF8"
        Write-Host "$inn-$iso_code-$salt-POS saving to C:\Temp\HostName.json."
        $hostname = "$inn-$iso_code-$salt-POS"   
    }
    

    $zabbixInstallFolder = "C:\Program Files\Zabbix Agent 2"
    Write-Host "Using $hostname in installation."
    msiexec.exe /l*v log.txt /i $zabbixAgentName /qn LOGTYPE=file LOGFILE=`"$zabbixInstallFolder\zabbix_agentd.log`" SERVER=office.retailservice24.ru SERVERACTIVE=office.retailservice24.ru HOSTNAME=$hostname TLSCONNECT=psk TLSACCEPT=psk TLSPSKIDENTITY=2839f5ebfd61d1ecf123be8ba458ed78 TLSPSKFILE=`"$zabbixInstallFolder\secret.psk`" TLSPSKVALUE=6a2a05db5cfa79cc1ffd6f9e18853140eeb36f306c8381be9bef1d8ebdec1cb6 HOSTMETADATA=$hostmetadata ENABLEPATH=1 INSTALLFOLDER=`"$zabbixInstallFolder`"
    Unregister-ScheduledTask -TaskName "InstallZabbixAgent" -Confirm:$false
}        
