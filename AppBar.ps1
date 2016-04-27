#First implementation of AppBar Class. <=Testing purpose
cls
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
cd $scriptPath

. '.\General.ps1'

#TODO: Split AppBar and AppBarIcon in two assemblies on disk, referencing each other
#TODO2:Put Region Designer to PowerShell.
AddAssembly "AppBar"

[Shell4Core.AppBar]$Bar = New-Object Shell4Core.AppBar("ABE_TOP", 0) -Property @{
	ClientSize = New-Object System.Drawing.Size(1940, 70)
	FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	Name = "AppBar"
	StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
}

[System.Windows.Forms.Button]$StartButton = New-Object System.Windows.Forms.Button -Property @{
	Location = New-Object System.Drawing.Point(10, 10)
	Name = "StartButton"
	Text = "Start"
	Size = New-Object System.Drawing.Size(50, 50)
	TabIndex = 1
}

[System.Windows.Forms.Panel]$OpenWindows = New-Object System.Windows.Forms.Panel -Property @{
	Anchor = 15 #All Edges
	Location = New-Object System.Drawing.Point(70, 10)
	Name = "RunningApps"
	Size = New-Object System.Drawing.Size(834, 50)
}

#TODO: Create DateLabel, group them in panel, center them.
[System.Windows.Forms.Label]$TimeLabel = New-Object System.Windows.Forms.Label -Property @{
	Text = [System.DateTime]::Now.ToString("HH:mm") + "
	" + [System.DateTime]::Today.ToString("dd.MM.yyyy")
	Size = New-Object System.Drawing.Size(95, 50)
	Location = New-Object System.Drawing.Point(($Bar.width - 105), 10)
	Anchor = 9 #Top Right
	Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, [Byte]0)
	TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
}

$Bar.SuspendLayout()
$Bar.Controls.AddRange(($StartButton, $OpenWindows, $TimeLabel))
$Bar.ResumeLayout($False)
$Bar.PerformLayout()

. '.\startMenu.ps1'
[System.Windows.Forms.Form]$startM = startMenu($Bar)

$StartButton.Add_Click(
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
		#$Bar.ticked($sender, $eventArgs)
		$TimeLabel.Text = [System.DateTime]::Now.ToString("HH:mm") + "
	" + [System.DateTime]::Today.ToString("dd.MM.yyyy")
		$intLeft = 0
		$OpenWindows.SuspendLayout()
		$OpenWindows.Controls.Clear()
		
		
		foreach($Win in (Get-Process | where-object {$_.MainWindowHandle -ne 0 -And $_.MainWindowTitle -ne ""} | Select-Object mainWindowTitle, mainwindowhandle | Sort-Object -Property MainWindowHandle))
		{
			#Write-Host ($Bar.EnumCallBack($Win.MainWindowHandle, $Win.MainWindowTitle))
			$d = New-Object Shell4Core.AppBarIcon $Win.MainWindowHandle
			$d.Text = $Win.MainWindowTitle
			$d.ArrangeMe([ref]$intLeft)
			$OpenWindows.Controls.Add($d)
		}
		$OpenWindows.ResumeLayout()
	})


$Bar.Add_Load(
	{
		write-host "Loading AppBar.."
		$Bar.RegisterBar()
	})

[System.Windows.Forms.Application]::Run(($Bar))

#Later...
#http://www.codeproject.com/Tips/895840/Multi-Threaded-PowerShell-Cookbook
#Solve the pipeline issue on close <= https://social.technet.microsoft.com/Forums/office/en-US/573c1870-5e31-43a1-a863-3f2ebed418df/how-to-handle-close-event-of-powershell-window-if-user-clicks-on-closex-button?forum=winserverpowershell