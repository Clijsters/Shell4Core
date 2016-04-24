. '.\General.ps1'

Function startMenu([System.Windows.Forms.Form]$appBar)
{
    [System.Windows.Forms.Form]$menuForm = New-Object System.Windows.Forms.Form
	$menuForm.SuspendLayout()
	$menuForm.visible = $False
	$menuForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
	$menuForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$menuForm.BackColor = [System.Drawing.SystemColors]::ControlDark
	$menuForm.ClientSize = New-Object System.Drawing.Size(464, 483)
	$menuForm.ShowInTaskbar = $False
	$menuForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$menuForm.Name = "StartMenu"
	$menuForm.Text = $menuForm.Name
	$menuForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
	    
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
            #$menuForm.left = $appBar.left
            #$menuForm.bottom = $appBar.Top
        }
        ABE_BOTTOM
		{
            #3
            $menuForm.left = $appBar.left
            $menuForm.bottom = $appBar.Top
        }
        default {Write-Host "This should never ever happen."}
    }
	
	$sc = New-Object System.Windows.Forms.SplitContainer
	$btnCmd = New-Object System.Windows.Forms.Button
	$btnComputer = New-Object System.Windows.Forms.Button
	$btnShutdown = New-Object System.Windows.Forms.Button
	#$sc.beginInit()
	$sc.Dock = [System.Windows.Forms.DockStyle]::Fill
	$sc.Location = New-Object System.Drawing.Point(0, 0)
	$sc.Size = New-Object System.Drawing.Size(464, 483)
	$sc.SplitterDistance = 223
	$sc.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$sc.TabIndex = 0
	
	$sc.Panel1.Controls.Add($btnCmd)
	$sc.Panel1.Controls.Add($btnComputer)
	$sc.Panel2.Controls.Add($btnShutdown)
	$sc.Panel2.BackColor = [System.Drawing.SystemColors]::ControlDark

	$btnComputer.Location = New-Object System.Drawing.Point(12, 12)
	$btnComputer.Name = "btnComputer"
	$btnComputer.Size = New-Object System.Drawing.Size(208, 40)
	$btnComputer.TabIndex = 0
	$btnComputer.Text = "Computer"
	$btnComputer.UseVisualStyleBackColor = $True

	$btnShutdown.Location = New-Object System.Drawing.Point(49, 431)
	$btnShutdown.Name = "btnShutdown"
	$btnShutdown.Size = New-Object System.Drawing.Size(176, 40)
	$btnShutdown.TabIndex = 1
	$btnShutdown.Text = "Shutdown"
	$btnShutdown.UseVisualStyleBackColor = $True
	
	$btnCmd.Location = New-Object System.Drawing.Point(12, 58)
	$btnCmd.Name = "btnCmd"
	$btnCmd.Size = New-Object System.Drawing.Size(208, 40)
	$btnCmd.TabIndex = 1
	$btnCmd.Text = "CommandPrompt"
	$btnCmd.UseVisualStyleBackColor = $True
	
	$menuForm.Controls.Add($sc)
	#$sc.EndInit()
	$menuForm.ResumeLayout($False)
	
	Return $menuForm
}
