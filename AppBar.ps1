#First implementation of AppBar Class. <=Testing purpose
#[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

cls
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
cd $scriptPath

. '.\General.ps1'
. '.\startMenu.ps1'

#TODO: Split AppBar and AppBarIcon in two assemblies on disk, referencing each other
#TODO2:Put Region Designer to PowerShell.
AddAssembly "AppBar"

[int]$intLeft = 0
[System.Collections.Generic.List[Shell4Core.AppBarIcon]]$ArrBtns = New-Object System.Collections.Generic.List[Shell4Core.AppBarIcon]
[System.Collections.Generic.List[Shell4Core.AppBarIcon]]$ArrRef  = New-Object System.Collections.Generic.List[Shell4Core.AppBarIcon]

Function RefreshOpenWindows($ArrRef){
	$ArrBtns.Clear()
	
	[Shell4Core.AppBar]::EnumWindows({
		Param($hWnd,$b)
		$Builder = New-Object System.Text.StringBuilder("", 256)
		If ([Shell4Core.AppBar]::IsWindowVisible($hWnd)) {
			[Shell4Core.AppBar]::GetWindowText($hWnd, $Builder, 256)
			$str = $Builder.ToString()
			If (($str -ne "") -And ($Str -ne "Start")) {
				$d = New-Object Shell4Core.AppBarIcon $hWnd
				$d.Text = $str
				$ArrBtns.Add($d)|Out-Null
			}
		}
		Return $True
	}, 0)
	
	#TODO: Find a better solution
	$oldMap = [Array]($ArrRef|Foreach-Object {$_.IsRestored})
	$newMap = [Array]($ArrBtns|Foreach-Object {$_.IsRestored})
	
	if ((diff $ArrRef $ArrBtns) -Or (diff $oldMap $newMap)){
		Write-Host "Something has changed. Let's refresh our OpenWindows-List"
		#TODO: First Order by ProcessName, then hWnd
		ForEach($btn in ($ArrBtns | Sort-Object -Property PtrhWnd)) {
			$btn.ArrangeMe([ref]$intLeft)
		}
		$OpenWindows.SuspendLayout()
		$OpenWindows.Controls.Clear()
		$OpenWindows.Controls.AddRange($ArrBtns)
		$OpenWindows.ResumeLayout($False)
		$ArrRef.Clear()
		#TODO: Find a better way to sync them. It seems that = doesn't work...
		ForEach($btn in $ArrBtns){$ArrRef.Add($btn)}
	}
	Return $ArrRef
}

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
	Anchor = ([int][System.Windows.Forms.AnchorStyles]::Top) + ([int][System.Windows.Forms.AnchorStyles]::Right) + ([int][System.Windows.Forms.AnchorStyles]::Left) + ([int][System.Windows.Forms.AnchorStyles]::Bottom) #15 = All Edges
	Location = New-Object System.Drawing.Point(70, 10)
	Name = "RunningApps"
	Size = New-Object System.Drawing.Size(834, 50)
}

[System.Windows.Forms.Label]$TimeLabel = New-Object System.Windows.Forms.Label -Property @{
	Text = [System.DateTime]::Now.ToString("HH:mm") + "
	" + [System.DateTime]::Today.ToString("dd.MM.yyyy")
	Size = New-Object System.Drawing.Size(95, 50)
	Location = New-Object System.Drawing.Point(($Bar.width - 105), 10)
	Anchor = ([int][System.Windows.Forms.AnchorStyles]::Top) + ([int][System.Windows.Forms.AnchorStyles]::Right) #9 = Top Right
	Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, [Byte]0)
	TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
}

[System.Windows.Forms.Timer]$myTime = New-Object System.Windows.Forms.Timer -Property @{
	Interval = 850
	Enabled = $True
}

[System.Windows.Forms.Form]$startM = startMenu([ref]$Bar)
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

$Bar.Add_Load({$Bar.RegisterBar()})

$myTime.Add_Tick(
	{
		#$Bar.ticked($sender, $eventArgs)
		$TimeLabel.Text = [System.DateTime]::Now.ToString("HH:mm") + "
	" + [System.DateTime]::Today.ToString("dd.MM.yyyy")
		$intLeft = 0
		$ArrRef = RefreshOpenWindows($ArrRef)
	})
	
$startM.Add_Load(
	{
		$startM.Top = $Bar.Bottom
		$StartM.Left = $Bar.Left
	}
)

$Bar.SuspendLayout()
$Bar.Controls.AddRange(($StartButton, $OpenWindows, $TimeLabel))
$Bar.ResumeLayout($False)
$Bar.PerformLayout()

[System.Windows.Forms.Application]::Run(($Bar))

#Later...
#http://www.codeproject.com/Tips/895840/Multi-Threaded-PowerShell-Cookbook
#Solve the pipeline issue on close <= https://social.technet.microsoft.com/Forums/office/en-US/573c1870-5e31-43a1-a863-3f2ebed418df/how-to-handle-close-event-of-powershell-window-if-user-clicks-on-closex-button?forum=winserverpowershell