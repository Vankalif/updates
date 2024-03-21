Add-Type -AssemblyName System.Web.Extensions

$RootFolder = "C:\ProgramData\RS24"
$conf_catalog = "$RootFolder\Conf\Zabbix"
$gather_data = "$RootFolder\GatherData\Zabbix"
$bin_catalog = "$RootFolder\Binaries"

$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$jsonString = (Get-Content "$conf_catalog\DadataTokens.json" -Encoding "UTF8")
$dadataObj = $serializer.DeserializeObject($jsonString)

$token = $dadataObj.token
$secret = $dadataObj.secret

$jsonString = (Get-Content "$conf_catalog\DataExport.json" -Encoding "UTF8")
$flag = $serializer.DeserializeObject($jsonString)
if ($flag.DataExportSuccess -eq "True")
{   
    $jsonString = (Get-Content "$gather_data\KKMData.json" -Encoding "UTF8")
    $KKMData = $serializer.DeserializeObject($jsonString)

    if (-not (Test-Path -Path "$conf_catalog\DadataInfo.json"))
    {
        $addr = $KKMData[0].pos_address
        $addr = '["' + $addr + '"]'
        
        $whttp = New-Object -ComObject "WinHttp.WinHttpRequest.5.1"
        $whttp.Open("POST", "https://cleaner.dadata.ru/api/v1/clean/address", $false)
        $whttp.SetRequestHeader("Accept", "application/json")
        $whttp.SetRequestHeader("Content-Type", "application/json"); 
        $whttp.SetRequestHeader("Authorization", "Token $token") 
        $whttp.SetRequestHeader("X-Secret", "$secret")


        $whttp.Send($addr)
        $stream = New-Object -ComObject "ADODB.Stream"
        $stream.Open()
        $stream.Type = 1
        $stream.Write($whttp.ResponseBody)
        $stream.SaveToFile("$gather_data\DadataInfo.json", 2)
        $stream.Close()
        $whttp = $null
        
    }

    $response = (Get-Content "$gather_data\DadataInfo.json" -Encoding "UTF8")
    $response = $serializer.DeserializeObject($response)[0]
    $iso_code = $response.region_iso_code
    $g_lat = $response.geo_lat
    $g_lon = $response.geo_lon
    $jsonOutput = @{geo_lat="$g_lat"; geo_lon="$g_lon"}
    $jsonOutput = $serializer.Serialize($jsonOutput)
    Set-Content "$gather_data\pos_geo_data.json" -Value $jsonOutput -Encoding "UTF8"
    $salt = -join ((65..90) | Get-Random -Count 9 | ForEach-Object {[char]$_})
    $inn = $KKMData[0].inn.Trim()
    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding
    $hostmetadata = ([System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($iso_code)))).replace("-","").ToLower()
    
    if ((Get-WmiObject win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit")
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-amd64-openssl.msi"
    }
    else
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-i386-openssl.msi"
    }
    
    if (Test-Path -Path "$conf_catalog\HostName.json")
    {
        $jsonData = (Get-Content "$conf_catalog\HostName.json" -Encoding "UTF8")
        $hostname = $serializer.DeserializeObject($jsonData)    
    }

    if ($null -eq $hostname.ZabbixHostName)
    {
        $jsonOutput = @{ZabbixHostName="$inn-$iso_code-$salt-POS"}
        $jsonOutput = $serializer.Serialize($jsonOutput)
        Set-Content "$conf_catalog\HostName.json" -Value $jsonOutput -Encoding "UTF8"
        $hostname = "$inn-$iso_code-$salt-POS"
    }

    $zabbixInstallFolder = "C:\Program Files\Zabbix Agent 2"
    Set-Location $bin_catalog
    cmd /c @"
msiexec.exe /l* log.txt /i $zabbixAgentName /qn LOGTYPE=file LOGFILE=`"$zabbixInstallFolder\zabbix_agentd.log`" SERVER=office.retailservice24.ru SERVERACTIVE=office.retailservice24.ru HOSTNAME=$hostname TLSCONNECT=psk TLSACCEPT=psk TLSPSKIDENTITY=2839f5ebfd61d1ecf123be8ba458ed78 TLSPSKFILE=`"$zabbixInstallFolder\secret.psk`" TLSPSKVALUE=6a2a05db5cfa79cc1ffd6f9e18853140eeb36f306c8381be9bef1d8ebdec1cb6 HOSTMETADATA=$hostmetadata ENABLEPATH=1 INSTALLFOLDER=`"$zabbixInstallFolder`"
"@
    SCHTASKS /Delete /TN "InstallZabbixAgent" /F
}        
