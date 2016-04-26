. '.\General.ps1'

#To use it with your own empty type see: https://gist.github.com/Clijsters/b846790b925dbdd2012cefd1dc538cfb
Function startMenu([System.Windows.Forms.Form]$appBar)
{
	#This is needed for calculating Control.Sizes
	[int]$myWidth = 464
	[int]$myHeight = 486 #TODO: myHeight = Slot(Buttons.Count + 1)+3 - Because Buttons will be enumed by config, this will come later
	[int]$btnMargin = 12
	[int]$btnHeight = 30
	
	#Predefined Slot algorithm for perfect placements of startMenuItems
	Function Slot([Int]$intSlot)
	{
		If ($intSlot -gt 0)
		{
			return ($btnMargin + ($intSlot * ($btnHeight + $btnMargin)))
		} Else 
		{
			return $btnMargin
		}
	}
	
	[System.Windows.Forms.Form]$menuForm = New-Object System.Windows.Forms.Form -Property @{
		visible = $False
		AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
		AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
		BackColor = [System.Drawing.SystemColors]::ControlDark
		ClientSize = New-Object System.Drawing.Size($myWidth, $myHeight)
		ShowInTaskbar = $False
		FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
		Name = "StartMenu"
		StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	}
	
	$menuForm.SuspendLayout()
	switch ($appBar.ABE)
	{
		ABE_LEFT
		{
			#0
		}
		ABE_TOP
		{
			#1
			$menuForm.left = $appBar.left
			$menuForm.top = $appBar.Bottom
		}
		ABE_RIGHT
		{
			#2
		}
		ABE_BOTTOM
		{
			#3
			$menuForm.left = $appBar.left
			$menuForm.bottom = $appBar.Top
		}
		default {Write-Host "This should never ever happen."}
	}
	
	$sc = New-Object System.Windows.Forms.SplitContainer -Property @{
		TabIndex = 0
		Dock = [System.Windows.Forms.DockStyle]::Fill
		Location = New-Object System.Drawing.Point(0, 0)
		Size = $menuForm.Size
		SplitterDistance = $myWidth * 0.45
		BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	}
	
	#If this goes bigger: http://powershell.org/wp/2013/01/23/join-powershell-hash-tables/
	[hashtable]$btnProto=@{
		Size = New-Object System.Drawing.Size(($sc.SplitterDistance - 2 * $btnMargin), $btnHeight)
		UseVisualStyleBackColor = $True
	}
	$btnComputer = New-Object System.Windows.Forms.Button -Property ($btnProto+@{
		Name = "btnComputer"
		Text = "Computer"
		TabIndex = 0
		Location = New-Object System.Drawing.Point($btnMargin, (Slot(0)))
	})
	
	$btnCmd = New-Object System.Windows.Forms.Button -Property ($btnProto+@{
		Name = "btnCmd"
		Text = "CommandPrompt"
		TabIndex = 1
		Location = New-Object System.Drawing.Point($btnMargin, (Slot(1)))
	})
	
	$btndrei = New-Object System.Windows.Forms.Button -Property ($btnProto+@{
		Name = "btndrei"
		Text = "Explorer"
		TabIndex = 2
		Location = New-Object System.Drawing.Point($btnMargin, (Slot(2)))
	})
	
	$btnShutdown = New-Object System.Windows.Forms.Button -Property ($btnProto+@{
		Name = "btnShutdown"
		Text = "Shutdown"
		TabIndex = 0
		Location = New-Object System.Drawing.Point($btnMargin, 431)
	})
	#startM was declared in AppBar. Would be nice, if this ould also work without AppBar (and the inherited class)
	$btnCmd.Add_Click(
		{
			Start-Process "CMD"
			$startM.Hide()
		}
	)
	$btndrei.Add_Click({Start-Process "explorer";$startM.Hide();})
	
	#$sc.beginInit()
	$sc.Panel1.Controls.AddRange(@(
		$btnComputer,
		$btnCmd,
		$btndrei
	))
	
	$sc.Panel2.Controls.Add($btnShutdown)
	$sc.Panel2.BackColor = [System.Drawing.SystemColors]::ControlDark
	#$sc.EndInit()
	
	$menuForm.Controls.Add($sc)
	$menuForm.ResumeLayout($False)
	
	Return $menuForm
}
