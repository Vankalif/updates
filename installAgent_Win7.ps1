Add-Type -AssemblyName System.Web.Extensions

$tmp_folder = "C:\Temp"
$prog_data = "C:\ProgramData"
set-location $tmp_folder

$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$jsonString = (Get-Content "$tmp_folder\DadataTokens.json" -Encoding "UTF8")
$dadataObj = $serializer.DeserializeObject($jsonString)

$token = $dadataObj.token
$secret = $dadataObj.secret

$jsonString = (Get-Content "$tmp_folder\DataExport.json" -Encoding "UTF8")
$flag = $serializer.DeserializeObject($jsonString)
if ($flag.DataExportSuccess -eq "True")
{
    
    $jsonString = (Get-Content "$prog_data\KKMData.json" -Encoding "UTF8")
    $KKMData = $serializer.DeserializeObject($jsonString)
    $addr = $KKMData[0].pos_address
    $addr = '["' + $addr + '"]'
    Set-Content "$tmp_folder\PosAddr.json" -Value $addr -Encoding "UTF8"
    
    $Accept = 'Accept: application/json'
    $Cont = 'Content-Type: application/json'
    $Auth = 'Authorization: Token ' + $token
    $XSec = 'X-Secret: ' + $secret
    cmd.exe /c @"
C:\Temp\curl-8.6.0_1-win32-mingw\bin\curl.exe -X POST -L "https://cleaner.dadata.ru/api/v1/clean/address" -H "$Cont" -H "$Accept" -H "$Auth" -H "$XSec" -d @PosAddr.json -o "response.json"
"@   
    $response = (Get-Content "$tmp_folder\response.json" -Encoding "UTF8")
    $response = $serializer.DeserializeObject($response)[0]
    
    $iso_code = $response.region_iso_code
    $g_lat = $response.geo_lat
    $g_lon = $response.geo_lon
    $jsonOutput = @{geo_lat="$g_lat"; geo_lon="$g_lon"}
    $jsonOutput = $serializer.Serialize($jsonOutput)
    Set-Content "$prog_data\POS_GEO_DATA.json" -Value $jsonOutput -Encoding "UTF8"
    
    $salt = -join ((65..90) | Get-Random -Count 9 | ForEach-Object {[char]$_})
    $inn = $KKMData[0].inn.Trim()
    
    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding
    $hostmetadata = ([System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($iso_code)))).replace("-","").ToLower()
    
    if ([Environment]::Is64BitOperatingSystem)
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-amd64-openssl.msi"
    }else
    {
        $zabbixAgentName = "zabbix_agent2-6.4.9-windows-i386-openssl.msi"
    }
    
    if (Test-Path -Path "$tmp_folder\HostName.json")
    {
        $jsonData = (Get-Content "$tmp_folder\HostName.json" -Encoding "UTF8")
        $hostname = $serializer.DeserializeObject($jsonData).ZabbixHostName    
    }

    if ($null -eq $hostname.ZabbixHostName)
    {
        $jsonOutput = @{ZabbixHostName="$inn-$iso_code-$salt-POS"}
        $jsonOutput = $serializer.Serialize($jsonOutput)
        Set-Content "$tmp_folder\HostName.json" -Value $jsonOutput -Encoding "UTF8"
        $hostname = "$inn-$iso_code-$salt-POS"
    }

    $zabbixInstallFolder = "C:\Program Files\Zabbix Agent 2"
    cmd /c @"
msiexec.exe /l* log.txt /i $zabbixAgentName /qn LOGTYPE=file LOGFILE=`"$zabbixInstallFolder\zabbix_agentd.log`" SERVER=office.retailservice24.ru SERVERACTIVE=office.retailservice24.ru HOSTNAME=$hostname TLSCONNECT=psk TLSACCEPT=psk TLSPSKIDENTITY=2839f5ebfd61d1ecf123be8ba458ed78 TLSPSKFILE=`"$zabbixInstallFolder\secret.psk`" TLSPSKVALUE=6a2a05db5cfa79cc1ffd6f9e18853140eeb36f306c8381be9bef1d8ebdec1cb6 HOSTMETADATA=$hostmetadata ENABLEPATH=1 INSTALLFOLDER=`"$zabbixInstallFolder`"
"@
    SCHTASKS /Delete /TN "InstallZabbixAgent" /F
}        
