#First implementation of AppBar Class. <=Testing purpose
cls
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
cd $scriptPath

. '.\General.ps1'

#TODO: Split AppBar and AppBarIcon in two assemblies on disk, referencing each other
#TODO2:Put Region Designer to PowerShell.
AddAssembly "AppBar"

$Bar = New-Object AppBar ABE_TOP

. '.\startMenu.ps1'
$startM = startMenu($Bar)

$Bar.startButton.Add_Click(
    {
        If (-not ($startM.Visible)) {
            $startM.Show()
        }
        else {
            $startM.Hide()
        }
    }
)

[System.Windows.Forms.Application]::Run(($Bar))

#Later...
#http://www.codeproject.com/Tips/895840/Multi-Threaded-PowerShell-Cookbook
