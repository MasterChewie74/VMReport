Function Send-Report {
    $Message="These are the VM Reports that were generated on $(Get-Date).<br> These reports were sent from this server: $env:COMPUTERNAME"
    $Title = "VM Reports $(Get-Date)"
    $Post = "<P>Thank You, <br>Systems Administrator</P>"

    $MessageParameters = @{
        Subject = "VM Reports $(Get-Date)"
        Body = ConvertTo-Html -Title $Title -Body $Message -Post $Post | Out-String
        From = "SysAdmin@domain.com"
        To = "User.Name@domain.com"
        SmtpServer = "smtp.domain.com"
        Attachments = (Get-ChildItem -Path "G:\VMReports\VMReports*.zip" -File).FullName
    }

    Send-MailMessage @MessageParameters -BodyAsHtml
}

 Function Fix-String {
    Param(
        [string]$StringToClean
    )

    $CleanString = ""
    
    if ($StringToClean.Length -ge 1) {
        $CleanString = $StringToClean.Substring(0)
    }

    $CleanString
}

#----------Set Credentials----------#
$User = "domain\user.name"
$PasswordFile = "G:\VMReports\info.txt"
$KeyFile = "G:\VMReports\AES.key"
$key = Get-Content $KeyFile
$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)

Remove-Item -Path "G:\VMReports\*.csv"
Move-Item -Path "G:\VMReports\*.zip" -Destination "G:\VMReports\Old"

#FQDN of vCenter servers
$vCenters = "vCenter1.domain.com","vCenter2.domain.com"

foreach ($vCenter in $vCenters) {
    $vCenterName = $vCenter.Split(".")[0]

    Set-PowerCLIConfiguration -Scope AllUsers -InvalidCertificateAction Ignore -ParticipateInCeip $false -Confirm:$false | Out-Null
    Connect-VIServer -Server $vCenter -Credential $Creds

    $AllVMs = Get-VM

    foreach ($VM in $AllVMs) {
        $VMIPs = Fix-String $VM.Guest.IPAddress
        
        $Props = @{
            'Name'=$VM.Name
            'Datacenter'=(Get-Datacenter -VM $VM.Name).Name
            'Power'=$VM.PowerState
            'Guest OS'=$VM.Guest.OSFullName
            'HardwareVersion'=$VM.HardwareVersion
            'IP Address'=$VMIPs
            'DNS Name'=$VM.ExtensionData.Guest.Hostname
            'CPUs'=$VM.NumCpu
            'CoresPerSocket'=$VM.CoresPerSocket
            'Memory Size'=$VM.MemoryGB
            'ProvisionedSpaceInGB'=[Math]::Round($VM.ProvisionedSpaceGB,0)
            'UsedSpaceInGB'=[Math]::Round($VM.UsedSpaceGB,0)
            'UUID'=$VM.PersistentId
            'HostCluster'=$VM.ResourcePool.Parent.Name
            'Host'=(Get-VMHost -VM $VM).Name.Split(".")[0].ToUpper()
            'Datastore'=(Get-Datastore -ID $VM.DatastoreIdList).Name
            'Managed By'=$VM.ExtensionData.Config.ManagedBy.ExtensionKey
            'Notes'=$Vm.Notes
        }

        $VMObject = New-Object -TypeName PSObject -Property $Props

        $VMObject | Select-Object -Property Name,Datacenter,Power,"Guest OS",HardwareVersion,"IP Address","DNS Name",CPUs,CoresPerSocket,"Memory Size",ProvisionedSpaceInGB,UsedSpaceInGB,HostCluster,Host,Datastore,"Managed By",Notes | Export-Csv -Path "G:\VMReports\$($vCenterName)-Unsorted.csv" -Append -NoTypeInformation


    }

    Import-Csv -Path "G:\VMReports\$($vCenterName)-Unsorted.csv" | Sort-Object Name | Export-Csv -Path "G:\VMReports\$($vCenterName)-VMs-$(Get-Date -f yyyy-MM-dd).csv" -NoTypeInformation
    Start-Sleep -Seconds 5
    Remove-Item -Path "G:\VMReports\$($vCenterName)-Unsorted.csv" -Force
    Remove-Item -Path G:\VMReports\ips.txt -Force
    Disconnect-VIServer -Server $vCenter -Confirm:$false
}

Compress-Archive -Path "G:\VMReports\*-VMs-*.csv" -DestinationPath "G:\VMReports\VMReports-$(Get-Date -f yyyy-MM-dd).zip"
Start-Sleep -Seconds 5
Send-Report
