<#
    .SYNOPSIS  
        Return Sentinel/Aladdin HASP Network Monitor metrics value, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return Sentinel/Aladdin HASP Network Monitor metrics value, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.2.1
        Name: Aladdin HASP Network Monitor Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 18MAR2016
        Testing environment: Windows Server 2008R2 SP1, Powershell 2.0, Aladdin HASP Network Monitor DLL 2.5.0.0 (hsmon.dll)

        Due _hsmon.dll_ compiled to 32-bit systems, you need to provide 32-bit environment to run all code, that use that DLL. You must use **32-bit instance of PowerShell** to avoid runtime errors while used on 64-bit systems. Its may be placed here:_%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe.

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get       - Get metric from collection's item;
            Count     - Count collection's items.
            DoCommand - Do NetHASP Monitor command that not require connection to server (HELP, VERSION). Command must be specified with -Key parameter

    .PARAMETER ObjectType
        Define rule to make collection:
            Server - NetHASP server (detected with "GET SERVERS" command);
            Slot - NetHASP key slot ("GET SLOTS ...");
            Module - NetHASP module ("GET MODULES ...");
            Login - authorized connects to NetHASP server ("GET LOGINS ...").

    .PARAMETER Key
        Define "path" to collection item's metric 

    .PARAMETER ServerID
        Used to select NetHASP server from list. 
        ServerID can be numeric (real ID) or alphanumeric (server name)
        Server name must be taked from field "NAME" of the "GET SERVERS" command output ('stuffserver.contoso.com' or similar).

    .PARAMETER ModuleID
        Used to additional objects selecting by Module Address

    .PARAMETER SlotID
        Used to additional objects selecting by Slot

    .PARAMETER LoginID
        Used to additional objects selecting by login Index

    .PARAMETER ErrorCode
        What must be returned if any process error will be reached

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        powershell -NoProfile -ExecutionPolicy "RemoteSigned" -File "nethasp.ps1" -Action "DoCommand" -Key "VERSION" -defaultConsoleWidth

        Description
        -----------  
        Get output of NetHASP Monitor VERSION command

    .EXAMPLE 
        ... "nethasp.ps1" -Action "Discovery" -ObjectType "Server" 

        Description
        -----------  
        Make Zabbix's LLD JSON for NetHASP servers

    .EXAMPLE 
        ... "nethasp.ps1" -Action "Get" -ObjectType "Slot" -Key "CURR" -ServerId "stuffserver.contoso.com" -SlotId "16" -ErrorCode "-127"

        Description
        -----------  
        Return number of used licenses on Slot #16 of stuffserver.contoso.com server. If processing error reached - return "-127"  

    .EXAMPLE 
        ... "nethasp.ps1" -Action "Get" -ObjectType "Module" -defaultConsoleWidth -Verbose

        Description
        -----------  
        Show formatted list of 'Module' object(s) metrics. Verbose messages is enabled. Console width is not changed.
#>

Param (
   [Parameter(Mandatory = $True)] 
   [ValidateSet('DoCommand', 'Discovery', 'Get', 'Count')]
   [String]$Action,
   [Parameter(Mandatory = $False)]
   [ValidateSet('Server', 'Module', 'Slot', 'Login')]
   [Alias('Object')]
   [String]$ObjectType,
   [Parameter(Mandatory = $False)]
   [String]$Key,
   [Parameter(Mandatory = $False)]
   [String]$ServerId,
   [Parameter(Mandatory = $False)]
   [String]$ModuleId,
   [Parameter(Mandatory = $False)]
   [String]$SlotId,
   [Parameter(Mandatory = $False)]
   [String]$LoginId,
   [Parameter(Mandatory = $False)]
   [String]$ErrorCode,
   [Parameter(Mandatory = $False)]
   [String]$ConsoleCP,
   [Parameter(Mandatory = $False)]
   [Switch]$DefaultConsoleWidth,
   [Parameter(Mandatory = $False)]
   [String]$HSMON_LIB_PATH,
   [Parameter(Mandatory = $False)]
   [String]$HSMON_INI_FILE,
   [Parameter(Mandatory = $False)]
   [Int]$HSMON_SCAN_TIMEOUT = 5
);

# Set default values from '1CHelper' module
if ( -not $HSMON_LIB_PATH ) {
    $HSMON_LIB_PATH = (Get-NetHaspDirectoryPath).Replace('\','\\')
}

if ( -not $HSMON_INI_FILE ) {
    $HSMON_INI_FILE = (Get-NetHaspIniFilePath).Replace('\','\\')
}

#Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
#[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US";

# Width of console to stop breaking JSON lines
if ( -not $CONSOLE_WIDTH ) {
    Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant;
}

# Full paths to hsmon.dll and nethasp.ini
if ( -not $HSMON_LIB_FILE ) {
    Set-Variable -Name "HSMON_LIB_FILE" -Value "$HSMON_LIB_PATH\\hsmon.dll" -Option Constant;
}
# Set-Variable -Name "HSMON_INI_FILE" -Value "$HSMON_LIB_PATH\\nethasp.ini" -Option Constant;

# Full path to hsmon.dll wrapper, that compiled by this script
if ( -not $WRAPPER_LIB_FILE ) {
    Set-Variable -Name "WRAPPER_LIB_FILE" -Value "$HSMON_LIB_PATH\\wraphsmon.dll" -Option Constant;
}

# Timeout in seconds for "SCAN SERVERS" connection stage
Set-Variable -Name "HSMON_SCAN_TIMEOUT" -Value $HSMON_SCAN_TIMEOUT # -Option Constant;

# Enumerate Objects. [int][NetHASPObjectType]::DumpType equal 0 due [int][NetHASPObjectType]::AnyNonexistItem equal 0 too
Add-Type -TypeDefinition "public enum NetHASPObjectType { DumpType, Server, Module, Slot, Login }";

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

Function Compile-WrapperDLL() {
   $WrapperSourceCode = 
@"
   using System;
   using System.Runtime.InteropServices;
   using System.Text;
   
   namespace HASP { 
      public class Monitor { 
         [DllImport(`"$($HSMON_LIB_FILE)`", CharSet = CharSet.Ansi,EntryPoint=`"mightyfunc`", CallingConvention=CallingConvention.Cdecl)]
         // String type used for request due .NET do auto conversion to Ansi char* with marshaliing procedure;
         // Byte[] type used for response due .NET char* is 2-byte, but mightyfunc() need to 1-byte Ansi char;
         // Int type used for responseBufferSize due .NET GetString() operate with [int] params. So, response lenght must be Int32 sized
         extern static unsafe void mightyfunc(string request, byte[] response, int *responseBufferSize);
     
         public Monitor() {}
      
         public static unsafe string doCmd(string request) {
            int responseBufferSize = 10240, responseLenght = 0;
            byte[] response = new byte[responseBufferSize];
            string returnValue = `"`";
            mightyfunc(request, response, &responseBufferSize);
            while (response[responseLenght++] != '\0') 
            returnValue = System.Text.Encoding.UTF8.GetString(response, 0, responseLenght);
            return returnValue;
         }
      } 
   }
"@

   $CompilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters;
   $CompilerParameters.CompilerOptions = "/unsafe /platform:x86";
   $CompilerParameters.OutputAssembly = $WRAPPER_LIB_FILE;
   Add-Type -TypeDefinition $WrapperSourceCode -Language CSharp -CompilerParameters $CompilerParameters;
   If ($False -Eq (Test-Path $WRAPPER_LIB_FILE)) {
      Write-Warning "Wrapper library not found after compilation. Something wrong";
      Exit;
   }
}

# Is this a Wow64 powershell host
Function Test-Wow64() {
    Return ((Test-Win32) -And (test-path env:\PROCESSOR_ARCHITEW6432))
}

# Is this a 64 bit process
Function Test-Win64() {
    Return ([IntPtr]::Size -Eq 8)
}

# Is this a 32 bit process
Function Test-Win32() {
    Return ([IntPtr]::Size -Eq 4)
}

Function Get-NetHASPData {
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [String]$Command,
      [Switch]$SkipScanning,
      [Switch]$ReturnPlainText
   );
   # Interoperation to NetHASP stages:
   #    1. Set configuration (point to .ini file)
   #    2. Scan servers while STATUS not OK or Timeout not be reached
   #    3. Do one or several GET* command   

   # Init connect to NetHASP module?   
   $Result = "";
   if (-Not $SkipScanning) {
      # Processing stage 1
      Write-Verbose "$(Get-Date) Stage #1. Initializing NetHASP monitor session"
      $Result = ([HASP.Monitor]::doCmd("SET CONFIG,FILENAME=$HSMON_INI_FILE")).Trim();
      if ('OK' -Ne $Result) { 
         Exit-WithMessage -Message "Error 'SET CONFIG' command: $Result" -ErrorCode $ErrorCode; 
      }
   
      # Processing stage 2
      Write-Verbose "$(Get-Date) Stage #2. Scan NetHASP servers"
      $Result = [HASP.Monitor]::doCmd("SCAN SERVERS");
      $ScanSeconds = 0;
      Do {
         # Wait a second before check process state
         Start-Sleep -seconds 1
         $ScanSeconds++; $Result = ([HASP.Monitor]::doCmd("STATUS")).Trim();
         #Write-Verbose "$(Get-Date) Status: $ret"
      } While (('OK' -ne $Result) -And ($ScanSeconds -Lt $HSMON_SCAN_TIMEOUT))

      # Scanning timeout :(
      If ($ScanSeconds -Eq $HSMON_SCAN_TIMEOUT) {
            Exit-WithMessage -Message "'SCAN SERVERS' command error: timeout reached" -ErrorCode $ErrorCode;
        }
    }

   # Processing stage 3
   Write-Verbose "$(Get-Date) Stage #3. Execute '$Command' command";
   $Result = ([HASP.Monitor]::doCmd($Command)).Trim();

   If ('EMPTY' -eq $Result) {
      Exit-WithMessage -Message "No data recieved" -ErrorCode $ErrorCode;
   } else {
      if ($ReturnPlainText) {
        # Return unparsed output 
        $Result;
      } else {
        # Parse output and push PSObjects to output
        # Remove double-quotes and processed lines that taking from splitted by CRLF NetHASP answer. 
        ForEach ($Line in ($Result -Replace "`"" -Split "`r`n" )) {
           If (!$Line) {Continue;}
           # For every non-empty line do additional splitting to Property & Value by ',' and add its to hashtable
           $Properties = @{};
           ForEach ($Item in ($Line -Split ",")) {
              $Property, $Value = $Item.Split('=');
              # "HS" subpart workaround
              if ($Null -Eq $Value) { $Value = "" }
              $Properties.$Property = $Value;
           } 
           # Return new PSObject with hashtable used as properties list
           New-Object PSObject -Property $Properties;
        }
      }
   }
}


####################################################################################################################################
#
#                                                 Main code block
#    
####################################################################################################################################
Write-Verbose "$(Get-Date) Checking runtime environment...";

# Script running into 32-bit environment?
If ($False -Eq (Test-Wow64)) {
   Write-Warning "You must run this script with 32-bit instance of Powershell, due wrapper interopt with 32-bit Windows Library";
   Write-Warning "Try to use %WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -File `"$HSMON_LIB_PATH\nethasp.ps1`" [-OtherOptions ...]";
   Exit;
}

Write-Verbose "$(Get-Date) Checking wrapper library for HASP Monitor availability...";
If ($False -eq (Test-Path $WRAPPER_LIB_FILE)) {
   Write-Verbose "$(Get-Date) Wrapper library not found, try compile it";
   Write-Verbose "$(Get-Date) First wrapper library loading can get a few time. Please wait...";
   Compile-WrapperDLL; 
#   [HASP.Monitor]::doCmd("VERSION");
} else {
  Write-Verbose "$(Get-Date) Loading wrapper library";
  Add-Type -Path $WRAPPER_LIB_FILE;
}

# Need to run one HASP command like HELP, VERSION ?
If ('DoCommand' -Eq $Action) {
   If ($Key) {
       Write-Verbose "$(Get-Date) Just do command '$Key'";
       ([HASP.Monitor]::doCmd($Key)).Trim();
   } Else {
      Exit-WithMessage -Message "No HASPMonitor command given with -Key option" -ErrorCode $ErrorCode;
   }
   Exit;
}

$Keys = $Key.Split(".");

# Exit if object is not [NetHASPObjectType]
#If (0 -Eq [int]($ObjectType -As [NetHASPObjectType])) { Exit-WithMessage -Message "Unknown object type: '$ObjectType'" -ErrorCode $ErrorCode; }

Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
# Object must contain Servers data?
if (($ObjectType -As [NetHASPObjectType]) -Ge [NetHASPObjectType]::Server) {
   Write-Verbose "$(Get-Date) Getting server list";
   $Servers = Get-NetHASPData -Command "GET SERVERS"; 
   if (-Not $Servers) { 
      Exit-WithMessage -Message "No NetHASP servers found" -ErrorCode $ErrorCode;
   }

   Write-Verbose "$(Get-Date) Checking server ID";
   if ($ServerId) {
      # Is Server Name into $ServerId
      if (![RegEx]::IsMatch($ServerId,'^\d+$')) {
         # Taking real ID if true
         Write-Verbose "$(Get-Date) ID ($ServerId) was not numeric - probaly its hostname, try to find ID in servers list";
         $ServerId = (PropertyEqualOrAny -InputObject $Servers -Property NAME -Value $ServerId).ID;
         if (!$ServerId) {
            Exit-WithMessage -Message "Server not found" -ErrorCode $ErrorCode;
         }
         Write-Verbose "$(Get-Date) Got real ID ($ServerId)";
      }
   }
   Write-Verbose "$(Get-Date) Filtering... (ID=$ServerId)";
   $Objects = $Servers = PropertyEqualOrAny -InputObject $Servers -Property ID -Value $ServerId
}

# Object must be processed with Servers data?
if (($ObjectType -As [NetHASPObjectType]) -ge [NetHASPObjectType]::Module) {
   Write-Verbose "$(Get-Date) Getting modules list"; 
   $Modules = ForEach ($Server in $Servers) { 
      Get-NetHASPData -Command "GET MODULES,ID=$($Server.ID)" -SkipScanning; 
   }
   $Objects = $Modules = PropertyEqualOrAny -InputObject $Modules -Property MA -Value $ModuleId
}

# Object must be processed with Servers+Modules data?
if (($ObjectType -As [NetHASPObjectType]) -ge [NetHASPObjectType]::Slot) {
   Write-Verbose "$(Get-Date) Getting slots list";
   $Slots = ForEach ($Module in $Modules) { 
      Get-NetHASPData -Command "GET SLOTS,ID=$($Module.ID),MA=$($Module.MA)" -SkipScanning; 
   }
   $Objects = $Slots = PropertyEqualOrAny -InputObject $Slots -Property SLOT -Value $SlotId
}

# Object must be processed with Servers+Modules+Slots data?
If (($ObjectType -As [NetHASPObjectType]) -Ge [NetHASPObjectType]::Login) {
   Write-Verbose "$(Get-Date) Getting logins list";
   # LOGININFO ignore INDEX param and return list of Logins anyway
   $Logins = ForEach ($Slot In $Slots) { 
      Get-NetHASPData -Command "GET LOGINS,ID=$($Slot.ID),MA=$($Slot.MA),SLOT=$($Slot.SLOT)" -SkipScanning;
   }
   $Objects = $Logins = PropertyEqualOrAny -InputObject $Slots -Property INDEX -Value $LoginId
}


ForEach ($Object in $Objects) {   
  Add-Member -InputObject $Object -MemberType NoteProperty -Name "ServerName" -Value (PropertyEqualOrAny -InputObject $Servers -Property ID -Value $Object.ID).Name;
  Add-Member -InputObject $Object -MemberType AliasProperty -Name "ServerID" -Value ID
}

Write-Verbose "$(Get-Date) Collection created, begin processing its with action: '$Action'";
switch ($Action) {
   'Discovery' {
      [Array]$ObjectProperties = @();
      Switch ($ObjectType) {
          'Server' {
             $ObjectProperties = @("SERVERNAME", "SERVERID");
          }
          'Module' {
             # MA - module address 
             $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "MAX");
          }
          'Slot'   {
             $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "SLOT", "MAX");
          }
          'Login' {
             $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "SLOT", "INDEX", "NAME");
          }
       }
       Write-Verbose "$(Get-Date) Generating LLD JSON";
       $Result =  Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
   }
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
      $Result = $(if ($Objects) { @($Objects).Count } else { 0 } ); 
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
