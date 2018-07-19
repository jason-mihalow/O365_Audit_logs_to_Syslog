<# Created by: Jason Mihalow

Description: This script utilizes a configured Azure application to gain access to the management API logs and send them to a logging server via TCP syslog.  This link will explain the prep work that needs to be done for this script to work: https://msdn.microsoft.com/en-us/office-365/get-started-with-office-365-management-apis.  This script is designed to pull only the Office 365 Exchange workload logs.  

Variables required to be configured:
$ClientID
$applicationID
$ClientSecret
$tenantdomain
#>

# This script will require the Web Application and permissions setup in Azure Active Directory
$ClientID       = ""             # Should be a ~35 character string insert your info here

#application ID
$applicationID  = ""

$ClientSecret   = ""       # Should be a ~44 character string insert your info here
$tenantdomain   = ""
$tenantID       = ""
$loginURL       = "https://login.windows.net"
$resource       = "https://manage.office.com"

#unformatted querytime
$querytime_unformatted = (get-date).AddMinutes(-10).ToUniversalTime()

#current time
$universaltime_current = (get-date).ToUniversalTime()

#format the querytime
$querytime = "{0:s}" -f $querytime_unformatted 

#destination logging server
$dstserver = ""

#destinaton logging port
$dstport = ""

#open TCP socket with logging server using .NET
$tcpConnection = New-Object System.Net.Sockets.TcpClient($dstserver, $dstport)
$tcpStream = $tcpConnection.GetStream()
$writer = New-Object System.IO.StreamWriter($tcpStream)
$writer.AutoFlush = $true

# Get an Oauth2 access token based on client id, secret and tenant domain
$body       = @{grant_type="client_credentials";resource=$resource;client_id=$applicationID;client_secret=$ClientSecret}

$oauth      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body

if ($oauth.access_token -ne $null) 
{
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

    $url = "https://manage.office.com/api/v1.0/$tenantID/activity/feed/subscriptions/content?contentType=Audit.Exchange&startTime=$querytime"
        
        Do 
        {
            $errorflag=$true
            Do
            {
                try
                {
                    $myReport = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $url)
                    $errorflag = $false
                }
                catch
                {
                    $Error
                    echo "here1"
                    start-sleep -s 5
                }
            }while($errorflag -eq $true)

            $timestamp_array = @()
            foreach ($event in ($myReport.Content | ConvertFrom-Json)) 
            {
                $contentUrl = $event.contentUri
                $errorflag = $true
                Do
                {
                    try
                    {
                        $myContentReport = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $contentUrl
                        $errorflag = $false
                    }
                    catch 
                    {
                        start-sleep -s 5
                    }
                 
                 }while($errorflag -eq $true)
                                
                foreach ($event in ($myContentReport.Content | ConvertFrom-Json))
                {
                    $line = ($event | Convertto-Json -Compress)
                    $line
                    $writer.Writeline($line)
                    $event.CreationTime
                    $timestamp_array += $event.CreationTime
                }
            }

            $url = $myReport.Headers.'NextPageUri'

        }while($url -ne $null)

    
    $timestamp_array = $timestamp_array | sort -Descending

    #the top array value is the latest timestamp; assign to timestamp type variable
    [datetime]$latest_stamp = $timestamp_array[0]

    $latest_stamp
 
    #find the differce in time between the latest_stamp and the current time; assign it to a variable
    $diff = New-TimeSpan -Start $latest_stamp -End $universaltime_current

    write-output ($diff | select Hours,Minutes,Seconds)

} else {

    Write-Host "ERROR: No Access Token"
}

