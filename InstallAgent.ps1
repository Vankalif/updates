$RootFolder = "C:\ProgramData\RS24"
$conf_catalog = "$RootFolder\Conf\Zabbix"
$gather_data = "$RootFolder\GatherData\Zabbix"
$bin_catalog = "$RootFolder\Binaries"

$DaData = (Get-Content "$conf_catalog\DadataTokens.json" -Encoding "UTF8") | ConvertFrom-Json
$token = $DaData.token
$secret = $DaData.secret

$flag = (Get-Content "$conf_catalog\DataExport.json" -Encoding "UTF8") | ConvertFrom-Json
if ($flag.DataExportSuccess -eq "True")
{
    $KKMData = Get-Content "$gather_data\KKMData.json" -Encoding "UTF8" | ConvertFrom-Json

    if (-not (Test-Path -Path "$conf_catalog\DadataInfo.json")) 
    {
        $addr = $KKMData[0].pos_address
        $headers = @{
            Accept = 'application/json; charset=utf-8'
            Authorization = 'Token ' + $token
            'X-Secret' = $secret
        }
        $body = ConvertTo-Json @("$addr")
        $response = Invoke-RestMethod 'https://cleaner.dadata.ru/api/v1/clean/address' -Method POST -ContentType "application/json; charset=Windows-1251" -Headers $headers -Body $body
        ConvertTo-Json -Compress -InputObject $response | Set-Content "$conf_catalog\DadataInfo.json" -Encoding "UTF8"
    }
    else
    {
        $response = Get-Content "$conf_catalog\DadataInfo.json" -Encoding "UTF8" | ConvertFrom-Json
        $response = $response.value[0]    
    }

    $iso_code = $response.region_iso_code
    $g_lat = $response.geo_lat
    $g_lon = $response.geo_lon
    @{geo_lat="$g_lat"; geo_lon="$g_lon"} | ConvertTo-Json | Set-Content "$gather_data\pos_geo_data.json" -Encoding "UTF8"
    $salt = -join ((65..90) | Get-Random -Count 9 | ForEach-Object {[char]$_})
    $inn = $KKMData[0].inn
    $inn = $inn.Trim()
    $mystream = [IO.MemoryStream]::new([byte[]][char[]]$iso_code)
    $hostmetadata = Get-FileHash -InputStream $mystream -Algorithm MD5
    $hostmetadata = $hostmetadata.hash.tolower()
    
    if ([Environment]::Is64BitOperatingSystem)
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-amd64-openssl.msi"
    }
    else
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-i386-openssl.msi"
    }

    if (Test-Path -Path "$conf_catalog\HostName.json")
    {
        $hostname = (Get-Content "$conf_catalog\HostName.json" -Encoding "UTF8") | ConvertFrom-Json
        $hostname = $hostname.ZabbixHostName
    }
    else
    {
        @{ZabbixHostName="$inn-$iso_code-$salt-POS"} | ConvertTo-Json | Set-Content "$conf_catalog\HostName.json" -Encoding "UTF8"
        $hostname = "$inn-$iso_code-$salt-POS"   
    }
    

    $zabbixInstallFolder = "C:\Program Files\Zabbix Agent 2"
    Set-Location $bin_catalog
    msiexec.exe /l*v log.txt /i $zabbixAgentName /qn LOGTYPE=file LOGFILE=`"$zabbixInstallFolder\zabbix_agentd.log`" SERVER=office.retailservice24.ru SERVERACTIVE=office.retailservice24.ru HOSTNAME=$hostname TLSCONNECT=psk TLSACCEPT=psk TLSPSKIDENTITY=2839f5ebfd61d1ecf123be8ba458ed78 TLSPSKFILE=`"$zabbixInstallFolder\secret.psk`" TLSPSKVALUE=6a2a05db5cfa79cc1ffd6f9e18853140eeb36f306c8381be9bef1d8ebdec1cb6 HOSTMETADATA=$hostmetadata ENABLEPATH=1 INSTALLFOLDER=`"$zabbixInstallFolder`"
    Unregister-ScheduledTask -TaskName "InstallZabbixAgent" -Confirm:$false
}        
