<#
.SYNOPSIS
    This script extracts information from netsh lease output.
 
.DESCRIPTION
    This script extracts information from netsh lease output. The netsh output is an array of strings, which makes it hard to programatically interact with the values.
    This script makes this interaction easier by converting the netsh output array of strings into PSCustomObjects with a property for each column in the netsh output.
    Only rows with valid MAC addresses are extracted.
 
.INPUTS
    DHCP server IP address, and the scope IP address.
 
.OUTPUTS
    A list of PSCustomObjects with the following headings: IpAddress, SubnetMask, MacAddress, LeaseExpiration, Type:
 
    Example output:
 
    IpAddress       : 10.19.10.8
    SubnetMask      : 255.255.254.0
    MacAddress      : 00-23-24-11-92-30
    LeaseExpiration : NEVER EXPIRES
    Type            : D
 
    IpAddress       : 10.19.10.11
    SubnetMask      : 255.255.254.0
    MacAddress      : 24-be-05-04-b5-a5
    LeaseExpiration : NEVER EXPIRES
    Type            : D
 
    IpAddress       : 10.19.11.254
    SubnetMask      : 255.255.254.0
    MacAddress      : 00-50-aa-26-a1-9c
    LeaseExpiration : NEVER EXPIRES
    Type            : D
 
 
.EXAMPLE
    $dhcpServerIpAddress = '10.19.10.1'
    $Scope = '10.19.10.0'
    $leasesCustomObjects = GetLeasesCustomObjects $dhcpServerIpAddress $Scope
 
.NOTES
    Author: dklempfner@gmail.com
    Date: 25/01/2017
#>
 
function GetLeasesCustomObjects
{
    Param([Parameter(Mandatory=$true)][String]$DhcpServerIpAddress,
          [Parameter(Mandatory=$true)][String]$Scope)
         
    $leases = ExtractLeasesFromNetshOutput $DhcpServerIpAddress $Scope
   
    $leaseCustomObjects = New-Object 'System.Collections.Generic.List[PSCustomObject]'
   
    foreach($lease in $leases)
    {
        if(!$lease -or $lease.Contains('0 in the Scope'))
        {
            continue
        }
        $ipAddress = ExtractIpAddressFromNetshOutputLine $lease
        $subnetMask = ExtractSubnetMaskFromNetshOutputLine $lease
        $macAddress = ExtractMacAddressFromNetshOutputLine $lease
        $leaseExpiration = ExtractLeaseExpirationFromNetshOutputLine $lease
        $type = ExtractTypeFromNetshOutputLine $lease
        $name = ExtractNameFromNetshOutputLine $lease
       
        $leaseCustomObject = [PSCustomObject]@{ IpAddress = $ipAddress; SubnetMask = $subnetMask; MacAddress = $macAddress; LeaseExpiration = $leaseExpiration; Type= $type; Name= $name}
        $leaseCustomObjects.Add($leaseCustomObject)
    }
   
    return $leaseCustomObjects
}
 
function ExtractLeasesFromNetshOutput
{
    Param([Parameter(Mandatory=$true)][String]$DhcpServerIpAddress,
          [Parameter(Mandatory=$true)][String]$Scope)
 
    $leasesInScopeNetshOutput = netsh dhcp server $DhcpServerIpAddress scope $Scope show clients 1   
    <#
    #Use this for testing:
    $leasesInScopeNetshOutput = 'Changed the current scope context to 10.19.10.0 scope.',
    '',
    'Type : N - NONE, D - DHCP B - BOOTP, U - UNSPECIFIED, R - RESERVATION IP',
    '============================================================================================',
    'IP Address      - Subnet Mask    - Unique ID           - Lease Expires        -Type -Name   ',
    '============================================================================================',
    '',
    '10.19.10.7      - 255.255.254.0  -07-0a-12-0a         - INACTIVE             -D-  BAD_ADDRESS',
    '10.19.10.9      - 255.255.254.0  -09-0a-12-0a         - INACTIVE             -D-  BAD_ADDRESS',
    '10.19.10.8      - 255.255.254.0  -00-23-24-11-92-30   - NEVER EXPIRES        -D-  ComputerOne',
    '10.19.10.11     - 255.255.254.0  -24-be-05-04-b5-a5   - NEVER EXPIRES        -D-  ComputerTwo',
    '10.19.11.254    - 255.255.254.0  -00-50-aa-26-a1-9c   - NEVER EXPIRES        -D-  ComputerThree',
    '',
    'No of Clients(version 4): 123 in the Scope : 10.19.10.0.',
    '',
    'Command completed successfully.'
    #>
    $startIndex = 9
    $numOfRowsBetweenLastIpAddressAndEndOfOutput = 5
    $endIndex = $leasesInScopeNetshOutput.Count - $numOfRowsBetweenLastIpAddressAndEndOfOutput
    $leases = $leasesInScopeNetshOutput[$startIndex..$endIndex]
   
    return $leases
}
 
function ExtractIpAddressFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)   
 
    $delimeter = '-'
 
    if(!$NetshOutputLine.Contains($delimeter))
    {
        Write-Error "Input line: $NetshOutputLine does not contain the delimeter $delimeter"
    }
 
    $indexOfFirstDash = $NetshOutputLine.IndexOf($delimeter)
   
    $ipAddress = $NetshOutputLine.Substring(0, $indexOfFirstDash).Trim()
    return $ipAddress
}
 
function ExtractSubnetMaskFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)   
 
    $delimeter = '-'
    $index = 2
   
    $firstDashIndex = $NetshOutputLine.IndexOf($delimeter)
    $secondDashIndex = GetNthOccurrenceOfCharsIndexInString $index $delimeter $NetshOutputLine
   
    $length = $secondDashIndex - $firstDashIndex - 1
   
    $subnetMask = $NetshOutputLine.Substring(($firstDashIndex + 1), $length).Trim()
    return $subnetMask
}
 
function ExtractMacAddressFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)   
        
    $macAddressRegex = [Regex]'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})'
   
    $macAddresses = $macAddressRegex.Matches($NetshOutputLine)
    if($macAddresses.Count -gt 1)
    {
        Write-Error 'Found more than one MAC address.'
    }
   
    $macAddress = $macAddresses[0].Value
   
    return $macAddress
}
 
function ExtractLeaseExpirationFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)   
    
    $startIndexOfLeaseExpiration = 56   
    
    $stringAfterStartIndexOfLeaseExpiration = $NetshOutputLine.Substring($startIndexOfLeaseExpiration)
   
    $firstIndexOfDash = $stringAfterStartIndexOfLeaseExpiration.IndexOf('-')
   
    $leaseExpiration = $stringAfterStartIndexOfLeaseExpiration.Substring(0, $firstIndexOfDash).Trim()
   
    return $leaseExpiration
}
 
function ExtractTypeFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)   
    
    $typeRegex = [Regex]'(-[A-Z]-)'
    $types = $typeRegex.Matches($NetshOutputLine)
   
    if($types.Count -gt 1)
    {
        Write-Error 'Found more than one Type.'
    }
   
    $type = $types.Value.Substring(1, 1)
    return $type
}
 
function ExtractNameFromNetshOutputLine
{
    Param([Parameter(Mandatory=$true)][String]$NetshOutputLine)
   
    $startIndex = ($NetshOutputLine.LastIndexOf('-') + 1)
    $length = $NetshOutputLine.Length - $startIndex
    $name = $NetshOutputLine.Substring($startIndex, $length).Trim()
    return $name
}
 
function GetNthOccurrenceOfCharsIndexInString($n, $char, $string)
{
    $errorIndicator = -1
    if ($string.Contains($char))
    {
        $stringSpitByChar = $string.Split($char)
       
        if($n -lt $stringSpitByChar.Count)
        {
            $index = 0
            for($i = 0; $i -lt $n; $i++)
            {
                if($i -eq 0)
                {
                    $index += $stringSpitByChar[$i].Length
                }
                else
                {
                    $index += $stringSpitByChar[$i].Length + 1
                }
            }
        }
        else
        {
            return $errorIndicator
        }
        return $index
    }
    return $errorIndicator
}