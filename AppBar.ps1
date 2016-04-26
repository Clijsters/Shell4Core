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

[System.Windows.Forms.Timer]$myTime = New-Object System.Windows.Forms.Timer -Property @{
	Interval = 2000
	Enabled = $True
}
$myTime.Add_Tick(
	{
		$Bar.ticked($sender, $eventArgs)
		$Bar.TimeLabel.Text = [System.DateTime]::Now.ToString("HH:mm")
	})
#Solve the pipeline issue on close
$myTime.Interval = 2000
$myTime.Enabled = $True

$Bar.Add_Load(
	{
		$Bar.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
		write-host "Loading AppBar.."
		$Bar.RegisterBar()
	})

[System.Windows.Forms.Application]::Run(($Bar))

#Later...
#http://www.codeproject.com/Tips/895840/Multi-Threaded-PowerShell-Cookbook
