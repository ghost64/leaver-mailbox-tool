#Generated Form Function
function GenerateForm {
    $testing = 0
    $version = 38
    $updatehost = "Update_Hostname"
    ########################################################################
    # Script Name: Leaver Mailbox Tool
    # Created: 27/05/2013
    # Created by: Michael Corrigan
    ########################################################################

    ########################################################################
    # TODO:
    # Give SD access
    # Test Refactor
    # Test & Fix backup
    # Check User serach (Move to disabled users lookup (QAD)
    # Try/Catch for:
    #       Mailbox Permissions
    #       Failure Emails
    ########################################################################

    #Import the Assemblies
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
    Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;

    #Creates all GUI Objects
    $exportForm = New-Object System.Windows.Forms.Form
    $userListBox = New-Object System.Windows.Forms.CheckedListBox
    $failedObjectCB = New-Object System.Windows.Forms.CheckBox
    $DisconnectRadioButton = New-Object System.Windows.Forms.RadioButton
    $exportRadioButton = New-Object System.Windows.Forms.RadioButton
    $bothRadioButton = New-Object System.Windows.Forms.RadioButton
    $startButton = New-Object System.Windows.Forms.Button
    $searchBox = New-Object System.Windows.Forms.TextBox
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    if ($env:USERNAME -like "a.*") { $to = $env:USERNAME.TrimStart("a.") ; $to = "$to@domainname.com"}else {$to = get-Mailbox $env:USERNAME | Select PrimarySmtpAddress; $to = [String]$to.PrimarySmtpAddress }
    $from = "From_Address"
    $smtpserver = "SMTP_Server"
    $adminbcc = "Admin_Email"
    $PSTArchive = "Archive_Location"
    $updatePath = "Update_Path"
    #Create empty vars
    $global:Users = @()
    $arrCheckedItems = @()
    $ArrayTempChecked = @()
    $global:CheckedResults = @()
    $SearchArrayTemp = @()
    $name = ""
    $newline = ""

    #Start main thread
    $handler_form1_Load= {
        mainTask
    }

    #Reloads User List with only users that contain what's in the search box
    function reloadUserList($searchStringTemp){
        $ArrayTempChecked = @()
        #Keeps the checked users
        $ArrayTempChecked +=  $userListBox.Items | ?{$userListBox.CheckedItems -contains $_}
        clearUserList
        $SearchArrayTemp = @()
        #Loops through and searches for the search term
        foreach ($UserTemp in $global:UsersExit) {
            if ($UserTemp -Match [string]$searchStringTemp) {
                $SearchArrayTemp += $UserTemp
            }
        }
        #Adds the searchlist to the userbox
        foreach ($UserSearch in $SearchArrayTemp){
            $userListBox.Items.add($UserSearch)
        }
        #Rechecking previous checked users
        foreach ($checkuser in $ArrayTempChecked){
            $i=0
            $SearchArrayTemp | foreach-object{
                foreach ($listuser in $_){
                    if ($listuser -eq $checkuser) {
                        $userListBox.SetItemChecked($i,$true)
                        $global:CheckedResults += $listuser
                    }
                    $i++
                }
            }
        }
    }

    #Clears User List
    function clearUserList(){
        $arrAllItems = @()
        $userListBox.Items | foreach-object {  $arrAllItems += $_}
        foreach($persontoremove in $arrAllItems){
            $userListBox.Items.remove($persontoremove)
        }
    }

    #Removes "Search..." after click
    $handler_SearchBox_Click={
        if ($searchBox.text -eq "Search...") {
            $searchBox.text = "";
        }
    }

    #Allows user to use the Return key to search
    $handler_SearchBox_KeyPress={
        #Event Argument: $_ = [System.Windows.Forms.KeyPressEventArgs]
        if ($_.KeyChar -eq 13) {
            $searchString = $searchBox.text
            reloadUserList $searchString
        }
    }

    #Reloads the local lmt file from the Exchange servers
    function reloadLocalCache () {
        $ServerList = "Mailbox_Server_1","Mailbox_Server_2"
        If (Test-Path ./users.lmt){Remove-Item ./users.lmt}
        #Let's start multi-threading this mofo
        foreach ($Server in $ServerList){
            start-job -scriptblock {param($server);
            Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
            Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
            get-Mailbox -ResultSize Unlimited -Server $Server | where {$_.Identity -like "*Disabled*"} | Select-Object PrimarySmtpAddress} -ArgumentList $Server
        }
        #Get,Wait,Receive from each thread
        $global:Users = Get-Job | Wait-Job | Receive-Job
        Remove-Job -State Completed
        #Add the users to the list
        $tempUserList = @()
        foreach ($User in $global:Users){
            $tempUser = $User.PrimarySmtpAddress -as [String]
            $tempUserList += $tempUser
        }
        $global:Users = $tempUserList | sort
        foreach ($User in $global:Users){
            $tempUser = $User
            $userListBox.Items.add($tempUser)
        }
        $global:UsersExit = @()
        $exportlistcount = $userListBox.Items.Count
        for ($index = 0; $index -lt $exportlistcount; ++$index){
            if($_.Index -ne $index){
                $global:UsersExit += $userListBox.Items[$index].ToString()
            }
        }
        #Export the list locally#
        $global:UsersExit | fl > ./users.lmt
    }

    function exportMailbox($listOfUserToWorkWith) {
        #if no users are selected
        if ([string]::IsNullOrEmpty($listOfUserToWorkWith)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a user or multiple users!" , "Error", 0)
        }
        else {
            #are you sure?
            if ($bothRadioButton.checked) {
                $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Do you want to export and disconnect the mailboxes tied to these addresses?`n$listOfUserToWorkWith" , "Status" , 4)
            }else {
                $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Do you want to export the mailboxes tied to these addresses?`n$listOfUserToWorkWith" , "Status" , 4)
            }
            if ($OUTPUT -eq "YES" ){
                #warns user that this could take forever
                [System.Windows.Forms.MessageBox]::Show("This may take some time" , "Warning", 0)
                #forced mode is off
                if (!$failedObjectCB.checked){
                    $arrCheckedItems | foreach-object{
                        $progressBar.Value = 0
                        $tempname = [string]$_
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $name =  Get-Mailbox $tempname | select Name | % {[string]$_.Name}
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $newline = Get-Mailbox $tempname | select Alias | % {[string]$_.Alias}
                        new-Item "$PSTArchive\$name" -type directory -Force
                        Add-MailboxPermission -identity $newline -accessrights fullaccess -user "$env:USERDOMAIN\$env:USERNAME"
                        $ArchivePSTName = "$PSTArchive\$name\$newline.pst"
                        $ArchiveFolderName = "$PSTArchive\$name"
                        do{Set-CasMailbox $newline -MAPIEnabled $True;[System.Windows.Forms.Application]::DoEvents()}
                        while(((get-CasMailbox $newline | select MAPIEnabled).MAPIEnabled -as [String]) -eq "False" )
                        New-MailboxExportRequest -Mailbox $newline -FilePath "$PSTArchive\$name\$newline.pst" -confirm:$false
                        $progressBar.Style = 2
                        $progressBar.MarqueeAnimationSpeed = 800;
                        $exportForm.refresh();
                        $exportForm.Text = "Leaver Mailbox Tool | Queueing..."
                        do{[System.Windows.Forms.Application]::DoEvents();Get-MailboxExportRequest -Status Queued | where { $_.Mailbox -like "*$name*" } | Resume-MailboxExportRequest;[System.Windows.Forms.Application]::DoEvents()}
                        while (![string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Queued | where { $_.Mailbox -like "*$name*" } )))
                        $progressBar.Style = 1
                        $progressBar.MarqueeAnimationSpeed = 0;
                        $exportForm.refresh();
                        $exportForm.Text = "Leaver Mailbox Tool | Exporting..."
                        do{if([string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } ))){$percent = Get-MailboxExportRequest | where { $_.Mailbox -like "*$name*" } | Get-MailboxExportRequestStatistics | select PercentComplete ; $percentComplete = $percent.PercentComplete; $progressBar.Value = $percentComplete;$exportForm.Text = "Leaver Mailbox Tool | Exporting... $percentComplete%";[System.Windows.Forms.Application]::DoEvents()}else {$logOutput = Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } | Get-MailboxExportRequestStatistics -IncludeReport | Format-List; Send-MailMessage -to $to -Bcc $adminbcc -subject "$tempname failed to export." -from $from  -body $logOutput  -smtpserver $smtpserver; Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } | Remove-MailboxExportRequest -confirm:$false; RemoveMailboxPermission; If (Test-Path $ArchivePSTName){Remove-Item $ArchivePSTName}; If (Test-Path $ArchiveFolderName){Remove-Item $ArchiveFolderName}; $progressBar.Value = 0; return}}
                        while ([string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Completed | where { $_.Mailbox -like "*$name*"} )))
                        Get-MailboxExportRequest -Status Completed | where { $_.Mailbox -like "*$name*" } | Remove-MailboxExportRequest -confirm:$false
                        Send-MailMessage -to $to -Bcc $adminbcc -subject "$tempname Exported." -from $from  -body "Mailbox successfully archived to $PSTArchive.  Location: <a href= '$PSTArchive\$name'>$PSTArchive\$name</a>" -BodyAsHTML -smtpserver $smtpserver
                        RemoveMailboxPermission
                        $progressBar.Value = 0
                    }
                    $exportForm.Text = "Leaver Mailbox Tool"
                }
                #forced mode is on
                elseif($failedObjectCB.checked){
                    $arrCheckedItems | foreach-object{
                        $progressBar.Value = 0
                        $tempname = [string]$_
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $name =  Get-Mailbox $tempname | select Name | % {[string]$_.Name}
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $newline = Get-Mailbox $tempname | select Alias | % {[string]$_.Alias}
                        new-Item "$PSTArchive\$name" -type directory -Force
                        Add-MailboxPermission -identity $newline -accessrights fullaccess -user "$env:USERDOMAIN\$env:USERNAME"
                        $ArchivePSTName = "$PSTArchive\$name\$newline.pst"
                        $ArchiveFolderName = "$PSTArchive\$name"
                        do{Set-CasMailbox $newline -MAPIEnabled $True;[System.Windows.Forms.Application]::DoEvents()}
                        while(((get-CasMailbox $newline | select MAPIEnabled).MAPIEnabled -as [String]) -eq "False" )
                        New-MailboxExportRequest -Mailbox $newline -FilePath "$PSTArchive\$name\$newline.pst" -BadItemLimit 49 -confirm:$false
                        $progressBar.Style = 2
                        $progressBar.MarqueeAnimationSpeed = 400;
                        $exportForm.refresh();
                        $exportForm.Text = "Leaver Mailbox Tool | Queueing..."
                        do{Get-MailboxExportRequest -Status Queued | where { $_.Mailbox -like "*$name*" }| Resume-MailboxExportRequest;[System.Windows.Forms.Application]::DoEvents()}
                        while (![string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Queued | where { $_.Mailbox -like "*$name*" })))
                        $progressBar.Style = 1
                        $progressBar.MarqueeAnimationSpeed = 0;
                        $exportForm.refresh();
                        $exportForm.Text = "Leaver Mailbox Tool | Exporting..."
                        do{if([string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } ))){$percent = Get-MailboxExportRequest | where { $_.Mailbox -like "*$name*" } | Get-MailboxExportRequestStatistics | select PercentComplete ; $percentComplete = $percent.PercentComplete; $progressBar.Value = $percentComplete;$exportForm.Text = "Leaver Mailbox Tool | Exporting... $percentComplete%";[System.Windows.Forms.Application]::DoEvents()}else {$logOutput = Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } | Get-MailboxExportRequestStatistics -IncludeReport | Format-List; Send-MailMessage -to $to -Bcc $adminbcc -subject "$tempname failed to export." -from $from  -body $logOutput  -smtpserver $smtpserver; Get-MailboxExportRequest -Status Failed | where { $_.Mailbox -like "*$name*" } | Remove-MailboxExportRequest -confirm:$false; RemoveMailboxPermission; If (Test-Path $ArchivePSTName){Remove-Item $ArchivePSTName}; If (Test-Path $ArchiveFolderName){Remove-Item $ArchiveFolderName}; $progressBar.Value = 0; return}}
                        while ([string]::IsNullOrEmpty((Get-MailboxExportRequest -Status Completed | where { $_.Mailbox -like "*$name*"} )))
                        Get-MailboxExportRequest -Status Completed | where { $_.Mailbox -like "*$name*" } | Remove-MailboxExportRequest -confirm:$false
                        Send-MailMessage -to $to -Bcc $adminbcc -subject "$tempname Exported." -from $from  -body "Mailbox successfully archived to $PSTArchive.  Location: <a href= '$PSTArchive\$name'>$PSTArchive\$name</a>" -BodyAsHTML -smtpserver $smtpserver
                        RemoveMailboxPermission
                        $progressBar.Value = 0
                    }
                    $exportForm.Text = "Leaver Mailbox Tool"
                }
            }
        }
    }

    function RemoveMailboxPermission {
        do{Remove-MailboxPermission -identity $newline -accessrights fullaccess -user "$env:USERDOMAIN\$env:USERNAME" -Confirm:$False}
        while ((Get-MailboxPermission $newline | where{$_.user -like "$env:USERDOMAIN\$env:USERNAME"} | select User) -ne $null)
    }

    function disconnectMailbox($listOfUserToWorkWith) {
        if ([string]::IsNullOrEmpty($listOfUserToWorkWith)) {
                [System.Windows.Forms.MessageBox]::Show("Please select a user or multiple users!" , "Error", 0)
            }
            else {
                if ($bothRadioButton.checked) {
                    $OUTPUT = "YES"
                }
                else {
                    $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Do you want to disconnect the mailboxes tied to these addresses?`n$listOfUserToWorkWith" , "Status" , 4)
                }
                if ($OUTPUT -eq "YES" ){
                    $exportForm.Text = "Leaver Mailbox Tool | Disconnecting.."
                    $arrCheckedItems | foreach-object{
                        $tempname = [string]$_
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $name =  Get-Mailbox $tempname | select Name | % {[string]$_.Name}
                        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                        $newline = Get-Mailbox $tempname | select Alias | % {[string]$_.Alias}
                        disable-mailbox -identity $newline -Confirm:$False
                        Send-MailMessage -to $to -Bcc $adminbcc -subject "$tempname disconnected." -from $from  -body "Mailbox successfully disconnected."  -smtpserver $smtpserver
                        $searchRemember = $searchBox.text
                        reloadUserList
                        $userListBox.Items.remove($_)
                        $global:UsersExit = @()
                        $exportlistcount = $userListBox.Items.Count
                        for ($index = 0; $index -lt $exportlistcount; ++$index){
                            if($_.Index -ne $index){
                                $global:UsersExit += $userListBox.Items[$index].ToString()
                            }
                        }
                        reloadUserList $searchRemember
                    }

                    $exportForm.Text = "Leaver Mailbox Tool"
                }

        }
    }

    #Go is clicked
    $handler_GO_Click={
        #magic testing mode
        if ($testing -ge 1) {
            $arrCheckedItems = @()
            $arrCheckedItems=$userListBox.Items | ?{$userListBox.CheckedItems -contains $_}
            $arrCheckedItems | foreach-object{
                $tempname = [string]$_
                Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                $name =  Get-Mailbox $tempname | select Name | % {[string]$_.Name}
                Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                $newline = Get-Mailbox $tempname | select Alias | % {[string]$_.Alias}
            }
        }
        else {
            $listOfUserToWorkWith = ""
            $arrCheckedItems = @()
            #get the checked users
            $arrCheckedItems=$userListBox.Items | ?{$userListBox.CheckedItems -contains $_}
            #check there are actually checked users
            if ($arrCheckedItems.length -gt 0) {
                #lists checked users name and mailbox size
                $arrCheckedItems | foreach-object{
                    $tempnameStat = $_ -as [string]
                    Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue;
                    #gets the users alias
                    $newlineStat = Get-Mailbox $tempnameStat | select Alias | % {[string]$_.Alias}
                    #gets the users mailbox size
                    $tempStatistics = Get-MailboxStatistics $tempnameStat | Select TotalItemSize
                    $listOfUserToWorkWith += "$_ `t"
                    $listOfUserToWorkWith += $tempStatistics.totalitemsize.value -as [string]
                    $listOfUserToWorkWith += "`n"
                }
            }
            #export is selected
            if ($exportRadioButton.checked) {
                exportMailbox $listOfUserToWorkWith
            }
            elseif ($DisconnectRadioButton.checked) {
                disconnectMailbox $listOfUserToWorkWith
            }
            elseif ($bothRadioButton.checked) {
                exportMailbox $listOfUserToWorkWith
                disconnectMailbox $listOfUserToWorkWith
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You have to select an option." , "Error!", 0)
            }
        }
    }

    function updateLMT {
        $updateversion = Get-Content "$updatePath\version"
        $testhost = hostname
        if($testhost -eq $updatehost){}
        if ($updateversion -gt $version) {
            [System.Windows.Forms.MessageBox]::Show("Updated needed. This is required. The LMT will close once finished launching." , "Update Found!", 0)
            Remove-Item -Path .\LeaverMailboxTool.ps1
            Copy-Item -Path $updatePath\LeaverMailboxTool.ps1 -Destination .\ -Force
            $updateFlag = 1
        }
        elseif ($updateversion -lt $version) {
                [System.Windows.Forms.MessageBox]::Show("This is a prerelease version" , "Beware of Dragons", 0)
        }
        else {
            return
        }
    }

    #Main thread
    function mainTask(){
        updateLMT
        Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Set-AdServerSettings -ViewEntireForest $True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($testing -gt 0 ) {
            [System.Windows.Forms.MessageBox]::Show("Testing mode is on.  You shouldn't be here." , "Beware of Dragons", 0)
        }
        if ([string]::IsNullOrEmpty((Get-ChildItem ./ | where { $_.Name -eq "users.lmt" }))){
            [System.Windows.Forms.MessageBox]::Show("Creating local cache.  This may take some time" , "Warning", 0)
            reloadLocalCache
        }
        else {
            $reloadOUTPUT= [System.Windows.Forms.MessageBox]::Show("Do you want to refresh local cache?" , "Leaver Mailbox Tool" , 4)
            if ($reloadOUTPUT -eq "YES") {
                [System.Windows.Forms.MessageBox]::Show("Refreshing local cache.  This may take some time" , "Warning", 0)
                reloadLocalCache
            }
            else {
                $global:UsersExit = Get-Content ./users.lmt
                $global:UsersExit = $global:UsersExit | sort
                foreach ($User in $global:UsersExit){
                    $tempUser = $User
                    $userListBox.Items.add($tempUser)
                }
            }
        }
        if ($updateFlag -ge 1) {
            $exportForm.close()
        }
    }

    $OnLoadForm_StateCorrection={
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $exportForm.WindowState = $InitialFormWindowState
    }

    #----------------------------------------------
    #region Generated Form Code
    $exportForm.AutoScaleMode = 3
    $exportForm.AutoSizeMode = 0
    $exportForm.AutoValidate = 2
    $exportForm.BackgroundImageLayout = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 431
    $System_Drawing_Size.Width = 581
    $exportForm.ClientSize = $System_Drawing_Size
    $exportForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $exportForm.FormBorderStyle = 1
    $exportForm.MaximizeBox = $False
    $exportForm.Name = "exportForm"
    $exportForm.Text = "Leaver Mailbox Tool"
    $exportForm.add_Load($handler_form1_Load)

    $userListBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $userListBox.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 12
    $System_Drawing_Point.Y = 12
    $userListBox.Location = $System_Drawing_Point
    $userListBox.Name = "userListBox"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 409
    $System_Drawing_Size.Width = 277
    $userListBox.Size = $System_Drawing_Size
    $userListBox.TabIndex = 12

    $exportForm.Controls.Add($userListBox)

    $failedObjectCB.DataBindings.DefaultDataSourceUpdateMode = 0

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 296
    $System_Drawing_Point.Y = 368
    $failedObjectCB.Location = $System_Drawing_Point
    $failedObjectCB.Name = "failedObjectCB"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 154
    $failedObjectCB.Size = $System_Drawing_Size
    $failedObjectCB.TabIndex = 11
    $failedObjectCB.TabStop = $True
    $failedObjectCB.Text = "Force!"
    $failedObjectCB.UseVisualStyleBackColor = $True

    $exportForm.Controls.Add($failedObjectCB)

    $DisconnectRadioButton.DataBindings.DefaultDataSourceUpdateMode = 0

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 296
    $System_Drawing_Point.Y = 308
    $DisconnectRadioButton.Location = $System_Drawing_Point
    $DisconnectRadioButton.Name = "DisconnectRadioButton"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 154
    $DisconnectRadioButton.Size = $System_Drawing_Size
    $DisconnectRadioButton.TabIndex = 10
    $DisconnectRadioButton.TabStop = $True
    $DisconnectRadioButton.Text = "Disconnect Mailbox Only"
    $DisconnectRadioButton.UseVisualStyleBackColor = $True

    $exportForm.Controls.Add($DisconnectRadioButton)


    $exportRadioButton.DataBindings.DefaultDataSourceUpdateMode = 0

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 296
    $System_Drawing_Point.Y = 278
    $exportRadioButton.Location = $System_Drawing_Point
    $exportRadioButton.Name = "exportRadioButton"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 154
    $exportRadioButton.Size = $System_Drawing_Size
    $exportRadioButton.TabIndex = 9
    $exportRadioButton.TabStop = $True
    $exportRadioButton.Text = "Export Mailbox Only"
    $exportRadioButton.UseVisualStyleBackColor = $True

    $exportForm.Controls.Add($exportRadioButton)

    $bothRadioButton.DataBindings.DefaultDataSourceUpdateMode = 0

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 296
    $System_Drawing_Point.Y = 338
    $bothRadioButton.Location = $System_Drawing_Point
    $bothRadioButton.Name = "bothRadioButton"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 24
    $System_Drawing_Size.Width = 180
    $bothRadioButton.Size = $System_Drawing_Size
    $bothRadioButton.TabIndex = 9
    $bothRadioButton.TabStop = $True
    $bothRadioButton.Text = "Export And Disconnect Mailbox"
    $bothRadioButton.UseVisualStyleBackColor = $True

    $exportForm.Controls.Add($bothRadioButton)

    $startButton.DataBindings.DefaultDataSourceUpdateMode = 0

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 497
    $System_Drawing_Point.Y = 369
    $startButton.Location = $System_Drawing_Point
    $startButton.Name = "startButton"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 23
    $System_Drawing_Size.Width = 75
    $startButton.Size = $System_Drawing_Size
    $startButton.TabIndex = 8
    $startButton.Text = "Go!"
    $startButton.UseVisualStyleBackColor = $True
    $startButton.add_Click($handler_GO_Click)

    $exportForm.Controls.Add($startButton)

    $searchBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 296
    $System_Drawing_Point.Y = 12
    $searchBox.Location = $System_Drawing_Point
    $searchBox.Name = "searchBox"
    $searchBox.Text = "Search..."
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 278
    $searchBox.Size = $System_Drawing_Size
    $searchBox.TabIndex = 3
    $searchBox.add_KeyPress($handler_SearchBox_KeyPress)
    $searchBox.add_Click($handler_SearchBox_Click)

    $exportForm.Controls.Add($searchBox)

    $progressBar.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 300
    $System_Drawing_Point.Y = 400
    $progressBar.Location = $System_Drawing_Point
    $progressBar.Name = "progressBar"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 270
    $progressBar.Size = $System_Drawing_Size
    $progressBar.Style = 1
    $progressBar.maximum = 100
    $progressBar.Step = 1
    $progressBar.TabIndex = 0

    $exportForm.Controls.Add($progressBar)

    #endregion Generated Form Code

    #Save the initial state of the form
    $InitialFormWindowState = $exportForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $exportForm.add_Load($OnLoadForm_StateCorrection)
    #Show the Form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $exportForm.ShowDialog()| Out-Null
} #End Function

#Call the Function
try {
    GenerateForm
}
finally {
    if ($global:UsersExit -ne $null) {
        If (Test-Path ./users.lmt){Remove-Item ./users.lmt}
        $global:UsersExit | fl > ./users.lmt
    }
}

