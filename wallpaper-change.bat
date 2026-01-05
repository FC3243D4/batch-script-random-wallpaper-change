<# : batch portion
:# The above line marks the beginning of a powershell comment block; and the Batch component of the Script. Do not modify.
::# Author: T3RRY ; Creation Date 12/02/2021 ; Version: 1.0.2
::# * Batch Powershell Hybrid * Resource: https://www.dostips.com/forum/viewtopic.php?f=3&t=5543
::# Script Purpose:
::# - Change the wallpaper from command prompt through the use of Parameter; Or by Input if no Parameter.
::# - Script Designed for use with pictures in the %Userprofile%\Pictures Directory
::#   or sub directories and should be placed in the %Userprofile%\Pictures Directory.
::#   - Hot tip: Add the %Userprofile%\Pictures Directory to your System environment Path variable.
::#     https://helpdeskgeek.com/windows-10/add-windows-path-environment-variable/

@Echo off & Mode 120,40

:# Test for Arg 1 ; Usage output ; Offer Input or Abort

 Set "Arg1=%~1"
 If "%Arg1%" == "" (
  Call "%~f0" "/?"
  Echo/
  Set "Wallpaper="
  Set /P "Wallpaper=Press ENTER to abort, or enter filepath / Search Term: "
  Setlocal EnableDelayedExpansion
  If "!Wallpaper!" == "" Exit /B
  Call "%~f0" "!Wallpaper!"
  Endlocal
  Exit /B
 )

:# Test for Unsupported Arg Count ; Notify Ignored Args; Show Help; Offer Abort

 Set ParamErr=%*
 If Not "%~2" == "" (
  Setlocal EnableDelayedExpansion
  Echo/Args:"!ParamErr:%Arg1% =!" Ignored. %~n0 only accepts 1 Arg.
  Call "%~f0" "/?"
  Endlocal
  Echo/Continue with Arg1:"%Arg1%" [Y]/[N]?
  For /F "Delims=" %%G in ('Choice /N /C:YN')Do if "%%G" == "N" Exit /B
 )
:# /Dir Switch - Display all image paths with matching extensions in tree of current Directory
 If Not "%Arg1:/Dir=%" == "%Arg1%" (
  Dir /B /S "*.jpg" "*.png" "*.bmp" | More
  Exit /B
 )

:# Usage test and output

 If Not "%Arg1:/?=%" == "%Arg1%" (
  Echo/ %~n0 Usage:
  Echo/
  Echo/ %~n0 ["wallpaper filepath" ^| "Search term"]
  Echo/      Search terms should include wildcard/s: * ? and / or extension as appropriate
  Echo/ Example:
  Echo/      Search for and select from .jpg files containing Dragon in the filename:
  Echo/     %~n0 "*Dragon*.jpg"
  Echo/
  Echo/ %~n0 [/Dir] - output list of available .jpg .png and .bmp files in the directory tree
  Echo/ %~n0 [/?] - help output
  Exit /B
 )

 Set "Wallpaper=%Arg1%"

:# Arg1 Not a valid path; Offer Addition of Wildcards to SearchTerm If not Present as Bookends
 If not exist "%Wallpaper%" If not "%Wallpaper:~0,1%" == "*" If not "%Wallpaper:~,-1%" == "*" (
  Echo/Add wildcards to "%Wallpaper%" {"*%Wallpaper%*"} [Y]/[N]?
  For /F "Delims=" %%G in ('Choice /N /C:YN')Do if "%%G" == "Y" Set "Wallpaper=*%Wallpaper%*"
 )

:# To support Search Terms run script in Top level of Directory containing Images; Find Full Path in Tree.

 PUSHD "%Userprofile%\Pictures"
 Set "Matches=0"
 (For /F "Delims=" %%G in ('Dir /B /S "%Wallpaper%"')Do (
   Set "Wallpaper=%%~fG"
   Set /A Matches+=1
   Call Set "Img[%%Matches%%]=%%~fG"
 )) 2> Nul

:# Determine if Target Wallpaper is Current Wallpaper; Notify and Exit
 reg query "HKEY_Current_User\Control Panel\desktop" -v wallpaper | %__AppDir__%findstr.exe /LC:"    wallpaper    REG_SZ    %Wallpaper%" && (
  Echo/Wallpaper already applied.
  Exit /B
 )

:# Enable environment for macro expansion, Arrays and code block variable operations
  Setlocal EnableExtensions EnableDelayedExpansion

 If NOT %Matches% GTR 1 Goto :Apply

:# Report When Multiple Matches found; Offer menu containing up to first 36 matches [ limit of menu macro ]

If %Matches% GTR 36 Set Matches=36

:# Menu Macro Author: T3RRY
:# Colored edition - Requires windows 10
:# IMPORTANT - RESERVED VARIABLES: Menu CH# CHCS Options Option Opt[i] Option.Output Cholist DIV

:# Menu macro escaped for Definition with DelayedExpansion Enabled. Ensures correct Environment:
 If not "!" == "" (
  Setlocal EnableExtensions EnableDelayedExpansion
 )

:# Test if virtual terminal codes enabled ; attempt to enable if false
 Reg Query HKCU\Console | %SystemRoot%\System32\findstr.exe /LIC:"VirtualTerminalLevel    REG_DWORD    0x1" > nul || (
   Reg Add HKCU\Console /f /v VirtualTerminalLevel /t REG_DWORD /d 1
 ) > Nul || (
  Echo(Virtual Terminal codes Required by Menu Macro not supported on your system.
  Pause
  Exit /B 1
 )

(Set \n=^^^

%= Newline var \n for multi-line macro definition - Do not modify. =%)

:# Key index list Allows maximum 36 menu options [ 0 indexed ]. Component of Menu Macro
 Set "ChoList=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

:# Get console width for dividing line
 for /F "usebackq tokens=2* delims=: " %%W in (`mode con ^| %SystemRoot%\System32\findstr.exe /LIC:"Columns"`) do Set /A "Console_Width=%%W"

:# Build Menu option color array [ RGB colormix ; red biased scaling darker. ]
 For /F %%E in ('Echo prompt $E^|cmd')Do Set "\E=%%E"
  Set "i=0"
 For /L %%i in (232 -4 70)Do (
  Set /A "BB=(%%i/4)+50,GG=%%i-BB,RR=GG+BB+(%%i/2)"
  Set "Color[!i!]=%\E%[38;2;!RR!;!GG!;!BB!m"
  Set /A i+=1
 )
 Set "i="

:# Define highlight color to use for a selected option
 Set "Highlight.Color=%\E%[0m%\E%[48;2;50;150;200m%\E%[31m%\E%[K"

:# Assign Flag true to enable Highlighting of selected option. Clears screen before menu is output.
 Set "Highlight=true"

:# Force console line dimensions to a size that supports highlight mode by preventing buffer scroll
 If /I "%Highlight%" == "true" (
  for /F "usebackq tokens=2* delims=: " %%W in (`mode con ^| %SystemRoot%\System32\findstr.exe /LIC:"Lines"`) do If %%W LSS 40 Mode !Console_Width!,40
 )

:# Build dividing line for menu output.
 Set "DIV=" & For /L %%i in (2 1 %Console_Width%)Do Set "DIV=!DIV!-"

:# Define dividing line Color
 Set "DIV=%\E%[33m%\E%[4m!DIV!%\E%[0m"

:# Menu macro Usage: %Menu% "quoted" "list of" "options"

     Set Menu=For %%n in (1 2)Do if %%n==2 (%\n%
%= Clear screen if highlight flag true  =%  If /I "^!Highlight^!" == "true" CLS%\n%
%= Output Dividing Line                 =%  Echo(^^!DIV^^!%\n%
%= Reset CH# index value for Opt[#]     =%  Set "CH#=0"%\n%
%= Undefine choice option key list      =%  Set "CHCS="%\n%
%= For Each in list;                    =%  For %%G in (^^!Options^^!)Do If not ^^!CH#^^! GTR 35 (%\n%
%= For Option Index value               =%   For %%i in (^^!CH#^^!)Do (%\n%
%= Build the Choice key list and Opt[#] =%    Set "CHCS=^!CHCS^!^!ChoList:~%%i,1^!"%\n%
%= array using the character at the     =%    Set "Opt[^!ChoList:~%%i,1^!]=%%~G"%\n%
%= current substring index.             =%    Set "option.output=%%~G"%\n%
%= Display ; removing # variable prefix =%    Echo(^^!Color[%%i]^^![^^!ChoList:~%%i,1^^!] ^^!Option.output:#=^^!%\E%[0m%\n%
%= Store line number of options         =%    Set "Line#^!ChoList:~%%i,1^!=%%i"%\n%
%= Increment Opt[#] Index var 'CH#'     =%    Set /A "CH#+=1"%\n%
%= Close CH# loop                       =%   )%\n%
%= Close Options loop                   =%  )%\n%
%= Output Dividing Line                 =%  Echo(^^!DIV^^!%\n%
%= Select option by character index     =%  For /F "Delims=" %%o in ('%SystemRoot%\System32\Choice.exe /N /C:^^!CHCS^^!')Do (%\n%
%= Assign return var 'OPTION' with the  =%   Set "Option=^!Opt[%%o]^!"%\n%
%= Highlight selected option with Line# =%   If /I "^!Highlight^!" == "true" (For /F "Delims=" %%X in ("^!Line#%%o^!")Do For /F "Delims=" %%Y in ('Set /A "%%X+2"')Do Echo(%\E%[%%Y;1H^^!Highlight.Color^^!{^^!Cholist:~%%X,1^^!} ^^!Opt[%%o]^^!%\E%[0m)%\n%
%= value selected from Opt[CH#] array.  =%   If /I "^!Option^!" == "Exit"   CLS ^& Exit /B 1%\n%
%= Exit type determines Errorlevel.     =%   If /I "^!Option^!" == "Previous" CLS ^& Exit /B 0%\n%
%= Return to previous script on Exit    =%  )%\n%
%= Move cursor to end of menu field     =%  If /I "^!Highlight^!" == "true" For /F "Delims=" %%Y in ('Set /A CH# + 2')Do Echo(%\E%[%%Y;1H%\n%
%= Capture Macro input - Options List   =% )Else Set Options=
========================================== :# End Definition of Menu Macro

:# Notify match count

 Echo/%Matches% Files Matched:"!Arg1!"

:# Use match count to build options list for Menu macro from IMG[ Array. Restricted to first 36 matches.
 
 Set "Menu.Options="
 For /L %%i in (1 1 %Matches%)Do Set "Menu.Options=!Menu.Options! "!Img[%%i]!""

:# Call Menu macro

 %Menu% %Menu.Options%

 Set "Wallpaper=%Option%"

 reg query "HKEY_Current_User\Control Panel\desktop" -v wallpaper | %__AppDir__%findstr.exe /LC:"    wallpaper    REG_SZ    %Wallpaper%" > nul && (
  Echo/Wallpaper already applied.
  Exit /B
 )

:Apply

:# Pipe Filepath to Powershell; Capture as Powershell Variable within Pipe; Exit on Return.
 Echo/!Wallpaper!| powershell.exe -noprofile "$Image = $input | ?{$_}; iex (${%~f0} | out-string)"
 Endlocal
 POPD
Exit /B 0

:# The below line Marks the end of a Powershell comment Block; And the End of the Batch Script. 
: end batch / begin powershell #>

<#
 Function Source: https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/
#>

Function Set-WallPaper {
 [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Image
    )

Add-Type -TypeDefinition @" 
using System; 
using System.Runtime.InteropServices;

public class Params
{ 
    [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
    public static extern int SystemParametersInfo (Int32 uAction, 
                                                   Int32 uParam, 
                                                   String lpvParam, 
                                                   Int32 fuWinIni);
}
"@ 

$SPI_SETDESKWALLPAPER = 0x0014
$UpdateIniFile = 0x01
$SendChangeEvent = 0x02

$RefreshIni = $UpdateIniFile -bor $SendChangeEvent

$ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $RefreshIni)
}

If (Test-Path $Image) {
    Set-WallPaper -Image $Image
    write-output "Wallpaper Updated."
}else {
    write-output "Wallpaper Does not exist in the Directory Tree."
}


