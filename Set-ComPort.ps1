param($remotePC, $oriCOM, $destCOM)

Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# Get available COM devices
Write-Host "Scanning for com port on" $remotePC
Write-Host "Wait please..."

[array]$COM = Get-WMIObject -ComputerName $remotePC Win32_PnPEntity | Where-Object {$_.Name -like "*(COM*"}
foreach ($port in $COM) {$port.Caption = $port.Name.Split("(")[1].Trim(")")} # Borrow Caption for COM-number

$COM = $COM | Sort-Object Caption

$COMlist = New-Object System.Collections.Specialized.OrderedDictionary # Use ordered list instead of messy hash table
foreach ($port in $COM) 
{
$COMlist.Add($port.Caption,$port.DeviceID)
Write-Host ($port.Caption + " " + $port.DeviceID)
}

# And then use i.e. $COMlist["COM2"] with the function.

function Set-ComPort 
{

Param ([string]$DeviceId, [ValidateScript({$_ -clike "COM*"})][string]$ComPort)

#Queries WMI for Device
$Device = Get-WMIObject Win32_PnPEntity | Where-Object {$_.DeviceID -eq $DeviceId}

#Execute only if device is present
if ($Device)
    {
    #Get current device info
    $DeviceKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $Device.DeviceID
    $PortKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $Device.DeviceID + "\Device Parameters"
    $Port = get-itemproperty -path $PortKey -Name PortName
    $OldPort = [convert]::ToInt32(($Port.PortName).Replace("COM",""))

    #Set new port and update Friendly Name
    $FriendlyName = $device.Name.split("(")[0] + "(" + $ComPort + ")"
    New-ItemProperty -Path $PortKey -Name "PortName" -PropertyType String -Value $ComPort -Force 
    #New-ItemProperty -Path $DeviceKey -Name "FriendlyName" -PropertyType String -Value $FriendlyName -Force 

    #Release Previous Com Port from ComDB
    $Byte = ($OldPort - ($OldPort % 8))/8
    $Bit = 8 - ($OldPort % 8)

    if ($Bit -eq 8) 
        {
        $Bit = 0 
        $Byte = $Byte - 1
        } 

    $ComDB = get-itemproperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\COM Name Arbiter" -Name ComDB
    $ComBinaryArray = ([convert]::ToString($ComDB.ComDB[$Byte],2)).ToCharArray()

    while ($ComBinaryArray.Length -ne 8) 
        {
        $ComBinaryArray = ,"0" + $ComBinaryArray
        } 

    $ComBinaryArray[$Bit] = "0"
    $ComBinary = [string]::Join("",$ComBinaryArray)
    $ComDB.ComDB[$Byte] = [convert]::ToInt32($ComBinary,2)
    Set-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\COM Name Arbiter" -Name ComDB -Value ([byte[]]$ComDB.ComDB) 

    } 

} 

# Thus, the usage for changing COM2 to COM8 would be

# Set-ComPort -DeviceID $COMlist["COM2"] -ComPort "COM8"

if ($COMlist[$oriCOM]){

    $password = ConvertTo-SecureString “acsopsL2” -AsPlainText -Force 
    $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList "rtladm" , $password
    
    Invoke-Command -ComputerName $remotePC -ScriptBlock ${Function:Set-ComPort} -ArgumentList $COMlist[$oriCOM] , $destCOM 

}
else { write-host $oriCOM " not found on " $remotePC
}
