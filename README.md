# Invoke-NoShell
Invoke-NoShell outputs a Mircosoft Office Word .doc file with an embedded macro.
It allows the automation of multiple similar versions of files, allowing to test how slight differences will effect it.
Currently, only Powershell payloads are supported.

The tool was launched at BSdiesTLV 2018, you may find the presentation in this repository.

## Parameters
Invoke-NoShell has 4 optional parameters:
```
-docPath [-D] Full path for the output file.
-payloadPath [-P] Full path of a text file containing Powershell payload to embed
-docGenerationMod [-M] A or M - for auto or manual payload generation mode
-lureText [-T] Text to be displayed in the doc for the unsuspecting victim
```

## Usage Example 
Create all 12 possible permutations with the lure text "open sesame" armed with the Powershell script 
```
Invoke-NoShell.ps1 -M A -T "Open sesame" -P c:\MyPowershellz\payload.ps1
```

Create a single document, manually select all the parameters
```
Invoke-NoShell.ps1 -M M
```

Create all 12 possible permutations in the folder C:\MyDocsFolder
```
Invoke-NoShell.ps1 -D C:\MyDocsFolder -M A
```

## Prerequisits
You need to have Microsoft Office installed in order to run this script.
The script will set the following key in order to allow automatic interaction with Word:
```
HKEY_CURRENT_USER\Software\Microsoft\Office\<OfficeVersion>\Word\Security\AccessVBOM
```
The tool was tested with Windows 10 x64, Office 2016 and PowerShell 5.0 but should work on other setups as well.

## //TODO:
Pull requests are welcomed:
+ One of the permutations is generated incorrectly at the moment, fixing it will require some refactoring.
+ Adding more features for generating the document, resulting in more permutations, for example - adding builtin obfuscation features.
+ Removing redundant functions written to the macro and never executed
+ <s>Alert on common VBA limits - lines longer than 1024 chars and non-ASCII chars.</s>

## References
Credit for Word COM object basics:  
https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/

HKCU execution policy bypass trick:  
https://blog.netspi.com/15-ways-to-bypass-the-powershell-execution-policy/

PowerShell ISE script loading documentation:  
https://docs.microsoft.com/en-us/powershell/scripting/core-powershell/ise/how-to-use-profiles-in-windows-powershell-ise
