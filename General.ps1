$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$Assemblies = @(
'System.Windows.Forms',
'System.Data',
'System.Drawing',
'System.Linq'
)

#2nd param - Assemblies. centralized it because unloading isn't possible
Function AddAssembly($TypeName)
{
	if (-not ([System.Management.Automation.PSTypeName]'$TypeName').Type)
	{
        $Code = [IO.File]::ReadAllText($scriptPath + "\" + $TypeName + ".vb")
        Write-Host "Compiling" $TypeName".vb and adding type..."
	    Add-Type -TypeDefinition $Code -ReferencedAssemblies $Assemblies -Language VisualBasic
    } else
	{
		Write-Host "Type already loaded."
	}
}

#The following is from https://gist.github.com/jakeballard/11240204
Function Set-WindowStyle {
param(
    [Parameter()]
    [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 
                 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 
                 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
    $Style = 'SHOW',
    
    [Parameter()]
    $MainWindowHandle = (Get-Process –id $pid).MainWindowHandle
)
    $WindowStates = @{
        'FORCEMINIMIZE'   = 11
        'HIDE'            = 0
        'MAXIMIZE'        = 3
        'MINIMIZE'        = 6
        'RESTORE'         = 9
        'SHOW'            = 5
        'SHOWDEFAULT'     = 10
        'SHOWMAXIMIZED'   = 3
        'SHOWMINIMIZED'   = 2
        'SHOWMINNOACTIVE' = 7
        'SHOWNA'          = 8
        'SHOWNOACTIVATE'  = 4
        'SHOWNORMAL'      = 1
    }
    
    $Win32ShowWindowAsync = Add-Type –memberDefinition @” 
[DllImport("user32.dll")] 
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru
    
    $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$Style]) | Out-Null
    Write-Verbose ("Set Window Style '{1} on '{0}'" -f $MainWindowHandle, $Style)
}
