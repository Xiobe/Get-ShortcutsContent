Function Get-ShortcutsContent{
 param(
   [Parameter(Mandatory=$true)]
   [string]$path_of_interest,

   [Parameter(Mandatory=$true)]
   [string]$extension
 )

 try {
   $Shortcuts = Get-ChildItem -Recurse $path_of_interest -Include $extension
 } catch {
   $Shell = New-Object -ComObject WScript.Shell
 }
 
 ForEach ($Shortcut in $Shortcuts)
 {
     $Properties = @{
     ShortcutName = $Shortcut.Name;
     ShortcutFull = $Shortcut.FullName;
     ShortcutPath = $shortcut.DirectoryName;
     Target       = $Shell.CreateShortcut($Shortcut).TargetPath
     }
   New-Object PSObject -Property $Properties
}
   [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
}
$Output = Get-ShortcutsContent -path_of_interest $env:HOMEDRIVE -extension *.lnk
$Output | Out-GridView ;Pause
