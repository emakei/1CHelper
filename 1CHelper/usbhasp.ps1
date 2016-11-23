<#
    .SYNOPSIS  
        Return USB (HASP) Device metrics value, count selected objects, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return USB (HASP) Device metrics value, count selected objects, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.2.1
        Name: USB HASP Keys Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 18MAR2016
        Testing environment: Windows Server 2008R2 SP1, USB/IP service, Powershell 2.0

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get       - Get metric from collection item;
            Count     - Count collection items.

    .PARAMETER ObjectType
        Define rule to make collection:
            USBController - "Physical" devices (USB Key)
            LogicalDevice - "Logical" devices (HASP Key)

    .PARAMETER Key
        Define "path" to collection item's metric 

    .PARAMETER PnPDeviceID
        Used to select only one item from collection

    .PARAMETER ErrorCode
        What must be returned if any process error will be reached

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        powershell.exe -NoProfile -ExecutionPolicy "RemoteSigned" -File "usbhasp.ps1" -Action "Discovery" -ObjectType "USBController"

        Description
        -----------  
        Make Zabbix's LLD JSON for USB keys

    .EXAMPLE 
        ... "usbhasp.ps1" -Action "Count" -ObjectType "LogicalDevice"

        Description
        -----------  
        Return number of HASP keys

    .EXAMPLE 
        ... "usbhasp.ps1" -Action "Get" -ObjectType "USBController" -PnPDeviceID "USB\VID_0529&PID_0001\1&79F5D87&0&01" -ErrorCode "-127" -DefaultConsoleWidth -Verbose

        Description
        -----------  
        Show formatted list of 'USBController' object metrics selected by PnPId "USB\VID_0529&PID_0001\1&79F5D87&0&01". 
        Return "-127" when processing error caused. Verbose messages is enabled. 

        Note that PNPDeviceID is unique for USB Key, ID - is not.
#>
Param (
   [Parameter(Mandatory = $False)] 
   [ValidateSet('Discovery','Get','Count')]
   [String]$Action,
   [Parameter(Mandatory = $False)]
   [ValidateSet('LogicalDevice','USBController')]
   [Alias('Object')]
   [String]$ObjectType,
   [Parameter(Mandatory = $False)]
   [String]$Key,
   [Parameter(Mandatory = $False)]
   [String]$PnPDeviceID,
   [Parameter(Mandatory = $False)]
   [String]$ErrorCode,
   [Parameter(Mandatory = $False)]
   [String]$ConsoleCP,
   [Parameter(Mandatory = $False)]
   [Switch]$DefaultConsoleWidth
);

#Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
# [System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

# Width of console to stop breaking JSON lines
if ( -not $CONSOLE_WIDTH ) {
    Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant
}

####################################################################################################################################
#
#                                                  Function block
#    
####################################################################################################################################
#
#  Select object with Property that equal Value if its given or with Any Property in another case
#
Function PropertyEqualOrAny {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [PSObject]$Property,
      [PSObject]$Value
   );
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         # IsNullorEmpty used because !$Value give a erong result with $Value = 0 (True).
         # But 0 may be right ID  
         If (($Object.$Property -Eq $Value) -Or ([string]::IsNullorEmpty($Value))) { $Object }
      }
   } 
}

#
#  Prepare string to using with Zabbix 
#
Function PrepareTo-Zabbix {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [String]$ErrorCode,
      [Switch]$NoEscape,
      [Switch]$JSONCompatible
   );
   Begin {
      # Add here more symbols to escaping if you need
      $EscapedSymbols = @('\', '"');
      $UnixEpoch = Get-Date -Date "01/01/1970";
   }
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         If ($Null -Eq $Object) {
           # Put empty string or $ErrorCode to output  
           If ($ErrorCode) { $ErrorCode } Else { "" }
           Continue;
         }
         # Need add doublequote around string for other objects when JSON compatible output requested?
         $DoQuote = $False;
         Switch (($Object.GetType()).FullName) {
            'System.Boolean'  { $Object = [int]$Object; }
            'System.DateTime' { $Object = (New-TimeSpan -Start $UnixEpoch -End $Object).TotalSeconds; }
            Default           { $DoQuote = $True; }
         }
         # Normalize String object
         $Object = $( If ($JSONCompatible) { $Object.ToString() } else { $Object | Out-String }).Trim();
         
         If (!$NoEscape) { 
            ForEach ($Symbol in $EscapedSymbols) { 
               $Object = $Object.Replace($Symbol, "\$Symbol");
            }
         }

         # Doublequote object if adherence to JSON standart requested
         If ($JSONCompatible -And $DoQuote) { 
            "`"$Object`"";
         } else {
            $Object;
         }
      }
   }
}

#
#  Convert incoming object's content to UTF-8
#
Function ConvertTo-Encoding ([String]$From, [String]$To){  
   Begin   {  
      $encFrom = [System.Text.Encoding]::GetEncoding($from)  
      $encTo = [System.Text.Encoding]::GetEncoding($to)  
   }  
   Process {  
      $bytes = $encTo.GetBytes($_)  
      $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
      $encTo.GetString($bytes)  
   }  
}

#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      # Init JSON-string $InObject
      $Result += "{$CRLF$Space`"data`":[$CRLF";
      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += "$Tab$Tab{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            If (-Not $itFirstProperty) { $Result += ",$Space" }
            $itFirstProperty = $False;
            $Result += "`"{#$Property}`":$Space$(PrepareTo-Zabbix -InputObject $Object.$Property -JSONCompatible)";
         }
         # No comma printed after last string
         $Result += "$Space}";
      }
   }
   End {
      # Finalize and return JSON
      "$Result$CRLF$Tab]$CRLF}";
   }
}

#
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Metric { 
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$Keys
   ); 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
        If ($Null -Eq $Object) { Continue; }
        # Expand all metrics related to keys contained in array step by step
        ForEach ($Key in $Keys) {              
           If ($Key) {
              $Object = Select-Object -InputObject $Object -ExpandProperty $Key -ErrorAction SilentlyContinue;
              If ($Error) { Break; }
           }
        }
        $Object;
      }
   }
}

#
#  Exit with specified ErrorCode or Warning message
#
Function Exit-WithMessage { 
   Param (
      [Parameter(Mandatory = $True, ValueFromPipeline = $True)] 
      [String]$Message, 
      [String]$ErrorCode 
   ); 
   If ($ErrorCode) { 
      $ErrorCode;
   } Else {
      Write-Warning ($Message);
   }
   Exit;
}

####################################################################################################################################
#
#                                                 Main code block
#    
####################################################################################################################################
# split key
$Keys = $Key.Split(".");

Write-Verbose "$(Get-Date) Taking Win32_USBControllerDevice collection with WMI"
$Objects = Get-WmiObject -Class "Win32_USBControllerDevice";

Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
Switch ($ObjectType) {
   'LogicalDevice' { 
      $PropertyToSelect = 'Dependent';    
   }
   'USBController' { 
      $PropertyToSelect = 'Antecedent';    
   }
}

# Need to take Unique items due Senintel used multiply logical devices linked to physical keys. 
# As a result - double "physical" device items into 'Antecedent' branch
#
# When the -InputObject parameter is used to submit a collection of items, Sort-Object receives one object that represents the collection.
# Because one object cannot be sorted, Sort-Object returns the entire collection unchanged.
# To sort objects, pipe them to Sort-Object.
# (C) PoSh manual
$Objects = $( ForEach ($Object In $Objects) { 
                 PropertyEqualOrAny -InputObject ([Wmi]$Object.$PropertyToSelect) -Property PnPDeviceID -Value $PnPDeviceID
           }) | Sort-Object -Property PnPDeviceID -Unique;

Write-Verbose "$(Get-Date) Processing collection with action: '$Action' ";
Switch ($Action) {
   # Discovery given object, make json for zabbix
   'Discovery' {
      Write-Verbose "$(Get-Date) Generating LLD JSON";
      $ObjectProperties = @("NAME", "PNPDEVICEID");
      $Result = Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
   }
   # Get metrics or metric list
   'Get' {
      If ($Keys) { 
         Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
         $Result = PrepareTo-Zabbix -InputObject (Get-Metric -InputObject $Objects -Keys $Keys) -ErrorCode $ErrorCode;
      } Else { 
         Write-Verbose "$(Get-Date) Getting metric list due metric's Key not specified";
         $Result = Out-String -InputObject $Objects;
      };
   }
   # Count selected objects
   'Count' { 
      Write-Verbose "$(Get-Date) Counting objects";  
      # if result not null, False or 0 - return .Count
      $Result = $( If ($Objects) { @($Objects).Count } Else { 0 } ); 
   }
}

# Convert string to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic chars for example)
If ($consoleCP) { 
   Write-Verbose "$(Get-Date) Converting output data to UTF-8";
   $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8; 
}

# Break lines on console output fix - buffer format to 255 chars width lines 
If (!$DefaultConsoleWidth) { 
   Write-Verbose "$(Get-Date) Changing console width to $CONSOLE_WIDTH";
   mode con cols=$CONSOLE_WIDTH; 
}

Write-Verbose "$(Get-Date) Finishing";
$Result;
