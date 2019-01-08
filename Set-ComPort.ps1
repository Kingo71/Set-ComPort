<#
.SYNOPSIS
Set-ComPort script by Angelo Lombardo @GRSD Amsterdam

.DESCRIPTION
Change COM port number on a remote pc 

.PARAMETER remotepc
HOST name or IP address of the remote PC
.PARAMETER oriCOM 
COM port to change (capital letter)
.PARAMETER  destCOM 
New COM port number (capital letter)
.EXAMPLE
./Set-ComPort.ps1 hostpc COM5 COM9  
#>
param($remotePC, $oriCOM, $destCOM)


# Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value "*" -Force


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


# Scanning for available COM devices on the remote machine
Write-Host "Scanning for com port on" $remotePC
Write-Host "Wait please..."

[array]$COM = Get-WMIObject -ComputerName $remotePC Win32_PnPEntity | Where-Object {$_.Name -like "*(COM*"}
foreach ($port in $COM) {$port.Caption = $port.Name.Split("(")[1].Trim(")")} #  Ecxtract Caption for COM number

$COM = $COM | Sort-Object Caption

$COMlist = New-Object System.Collections.Specialized.OrderedDictionary # Create an ordered list with COM/DEviceID pair

foreach ($port in $COM) 
{
$COMlist.Add($port.Caption,$port.DeviceID)
Write-Host ($port.Caption + " " + $port.DeviceID)
}

# Set-ComPort $COMlist["COM2"] , "COM8"  --  Usage example

# Check if COM port exits and invoke the function on the remote machine
if ($COMlist[$oriCOM]){

    Invoke-Command -ComputerName $remotePC -ScriptBlock ${Function:Set-ComPort} -ArgumentList $COMlist[$oriCOM] , $destCOM 

}
else { write-host $oriCOM " not found on " $remotePC
}

