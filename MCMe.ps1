write-host ""
write-host "44	44					" -ForegroundColor white
write-host "44	44	11		" -ForegroundColor white
write-host "44	44					" -ForegroundColor white
write-host "44	44	22		" -ForegroundColor white
write-host "4444444444	22		" -ForegroundColor white
write-host "44	44	22		" -ForegroundColor white
write-host "44	44	22		" -ForegroundColor white
write-host "44	44	22		" -ForegroundColor white
write-host "44	44	22		" -ForegroundColor white
write-host ""
write-host ""
write-host " Available commands are as followed" -ForegroundColor Blue
write-host ""
write-host "|--|ct365 - Connect to 365 (Exchange)"
write-host "|--|ctleavers - Connect to Security And Compliance Centre"
write-host ""
write-host ""
write-host "**|**singleMailbox - Returns details of a mailbox"
write-host "**|**addPermission - Adds mailbox Permission"
write-host "**|**getAllFolderInformation - Gets all mailbox folders information - (Useful when a user is unable to find a folder in their mailbox"
write-host "**|**getMailboxTotalSize - Returns a full size of the mailbox"
write-host "**|**oOfOffice - Sets up out of office"
write-host "**|**arrayWithValues - creates a dynamic array and users a foreach loop which is dependent upon the amount of entered elements"
write-host ""
write-host ""
write-host "#!---auditLog - Generates and saves the log to an mDrive [Requires a fix as M drives are no longer being provisioned]"
write-host "#!---usersToDistrGroup - Adds users to security group (works only on single users - That's how function is designed."
write-host "#!---getDdlContents - generates the list of the users for the DDL"
write-host "#!---checkForwardingSMTP - Checks for forwarding SMTP Address"
write-host ""
write-host ""
write-host "This is a powershell module created by Hubert Dziedziczak." -ForegroundColor yellow
write-host "Please use it to your own extent, and if you publish it somewhere please ensure to credit my work." -ForegroundColor yellow 
write-host "Enjoy managing 365 in much easier way! ;) ;) ;) ;) ;) ;) " -ForegroundColor yellow



#Current Modules
#!<-- The current functions within this module are as below:
# - ct365 - Connect to 365 (Exchange)
# - ctleavers - Connect to Security And Compliance Centre

#<!-Mailbox-!>

# - singleMailbox - Returns details of a mailbox
# - addPermission - Adds mailbox Permission
# - getAllFolderInformation - Gets all mailbox folders information - (Useful when a user is unable to find a folder in their mailbox
# - getMailboxTotalSize - Retursns a full size of the mailbox [ THIS requires an update. At the moment the username is static and it won't actually be dynamic against a different user]
# - oOfOffice - Sets up out of office.

#<!-AD & Other useful commands-!>

# - auditLog - Generates and saves the log to an mDrive [Requires "D:\Logs" area.]
# - usersToDistrGroup - Adds users to security group (works only on single users - That's how function is designed].
# - getDdlContents - generates the list of the users for the DDL


#Connects you with Exchange Server
function ct365
{
$uname = whoami /upn
	Connect-EXOPSSession -UserPrincipalName $uname
}

#Connects you with Compliance Search
function ctLeavers{
$uname = whoami /upn
	Connect-IPPSSession -UserPrincipalName $uname
}

#Returns details of a single Mailbox
function singleMailbox([string]$arg1){
    get-mailbox -identity $arg1 | fl
}

#Connects you with Compliance Search
$Button365_Click = 
{
[System.Windows.Forms.MessageBox]::Show("Connecting to 365, please enter your 2FA Code" , "2FA Authenticator")
$form.hide()
ct365
$form.show()
[System.Windows.Forms.MessageBox]::Show("Connected" , "2FA Authenticator")
}

$Button2_Click = 
{
[System.Windows.Forms.MessageBox]::Show("Connecting to Compliance Center, please enter your 2FA Code" , "2FA Authenticator")
$form.hide()
ctleavers
$form.show()
[System.Windows.Forms.MessageBox]::Show("Connected" , "2FA Authenticator")
}

$Button3_Click = 
{

[System.Windows.Forms.MessageBox]::Show("On the next screen, please enter the SMTP of the mailbox information you'd like to retrieve" , "Returning the mailbox")
$textbox.add_KeyDown({
if($_.KeyCode -eq "Enter"){
$textbox.text | out-Host}})
singleMailbox($._textbox.text)
$form.hide()
[System.Windows.Forms.MessageBox]::Show("For the Result, please see powershell screen" , "Returning the mailbox")
$form.show()
}

Function tool_Open{

    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing
	
    
    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Total Tool v2 developed by HDD"
    $Form.Size = New-Object System.Drawing.Size(800,400)
    $Form.StartPosition = "CenterScreen"
    $Form.Topmost = $True

    # Add Button_ct365
    $Button365 = New-Object System.Windows.Forms.Button
    $Button365.Location = New-Object System.Drawing.Size(35,35)
    $Button365.Size = New-Object System.Drawing.Size(120,23)
    $Button365.Text = "Connect to 365"
	
	# Add Button for Leavers
	$Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(175,35)
    $Button2.Size = New-Object System.Drawing.Size(120,23)
    $Button2.Text = "Connect to Compliance"
	
	$Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(315,35)
    $Button3.Size = New-Object System.Drawing.Size(120,23)
    $Button3.Text = "Get Mailbox"
	
	$textbox = New-Object System.Windows.Forms.TextBox

    $Form.Controls.Add($Button365)
	$Form.Controls.Add($Button2)
	$Form.Controls.Add($Button3)
	$form.Controls.add($textbox)

    #Add Button event 
    $Button365.Add_Click($Button365_Click)
	
	$Button2.add_Click($Button2_Click)
	
	$Button3.add_Click($Button3_Click)

         #Show the Form 
    $form.ShowDialog()| Out-Null 
 } #End Function 

#retrieves audit log and saves it to D:\Logs\....
function auditlog{
[string]$upn = Read-Host -prompt "What is the UPN on the mailbox you'd like to run AUDIT LOG on?"
Search-MailboxAuditLog -Identity $upn -LogonTypes Owner,Admin,Delegate -ShowDetails -ResultSize 50000 | Select -property LastAccessed,Operation,OperationResult,LogonType,DestFolderPathName,LogonUserDisplayName,SourceItemSubjectsList | export-csv "D:\Logs\RequestedAuditLog.csv"
}

#Adds mailbox permission
function addPermission{
[string]$upn = Read-Host -prompt "What is the targetted mailbox?"
[string]$username = read-host  -prompt  "What is the user you'd like to assign access to?"
add-mailboxPermission -identity $upn -user $username -accessrights fullaccess -automapping: $false -inheritancetype all
}

#adds user to security group
function usersToDistrGroup{
[string]$upn = Read-Host -prompt "What is the targetted distribution Group?"
[string]$username = read-host  -prompt  "What is the user you'd like to add?"
Write-Host "Adding name to distributionGroup"
Add-DistributionGroupMember -Identity $distributionGroup -Member $name -BypassSecurityGroupManagerCheck
}

#returns the contents of DDL. #requires to add check for if statement.
function getDdlContents{
[string]$ddl = Read-Host -prompt "What is the targetted DDL?"
$1 = get-dynamidDistributionGroup $ddl
get-recipient -recipientPreviewFilter $1.recipientFilter -resultsize Unlimited
}

#returns every mailbox folder information - Name, Path, Amount of items and foldersize
function getAllFolderInformation{
[string]$upn = Read-Host -prompt "Who's mailbox folders would you like to see?"
Get-MailboxFolderStatistics -Identity $upn -IncludeOldestAndNewestItems | format-table Name, FolderPath, ItemsInFolder, FolderSize, OldestItemReceivedDate, CreationTime}

#Returns the size of the mailbox.
function getMailboxTotalSize{
[string]$upn = Read-Host -prompt "Who's mailbox size would you like to see?"
$Mailboxes = Get-Mailbox -Identity $upn -ResultSize Unlimited | Select UserPrincipalName, Identity, ArchiveStatus

$count = 0

$MailboxSizes = @()

	foreach ($Mailbox in $Mailboxes){
		$count++

		$MailboxProperties = New-Object PSObject

		$MailboxStatistics = Get-MailboxStatistics $Mailbox.UserPrincipalName | Select LastLogonTime, TotalItemSize, ItemCount
        $MailboxProperties | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
        $MailboxProperties | Add-Member -MemberType NoteProperty -Name "Last Logon Time" -Value $MailboxStatistics.LastLogonTime
        $MailboxProperties | Add-Member -MemberType NoteProperty -Name "Mailbox Size" -Value $MailboxStatistics.TotalItemSize
        $MailboxProperties | Add-Member -MemberType NoteProperty -Name "Mailbox Item Count" -Value $MailboxStatistics.ItemCount

			if($Mailbox.ArchiveStatus -eq "Active"){
				$MailboxArchiveStatistics = Get-MailboxStatistics $Mailbox.UserPrincipalName -Archive | Select TotalItemSize, ItemCount
				$MailboxProperties | Add-Member -MemberType NoteProperty -Name "Archive Size" -Value $MailboxArchiveStatistics.TotalItemSize
				$MailboxProperties | Add-Member -MemberType NoteProperty -Name "Archive Item Count" -Value $MailboxArchiveStatistics.ItemCount
			}else{
				$MailboxProperties | Add-Member -MemberType NoteProperty -Name "Archive Size" -Value 0
				$MailboxProperties | Add-Member -MemberType NoteProperty -Name "Archive Item Count" -Value 0
			}

		if($count%10 -eq 0){
			Write-Host "$count Mailboxes have been processed"
				}else{
			Write-Host $count
			}

	$MailboxSizes += $MailboxProperties
	$MailboxProperties = $null
}

$MailboxSizes | export-csv D:\Logs -NoTypeInformation
}

#Out Of Office remotely without assigning myself access to someone else's mailbox.
function oOfOffice{
[string]$upn = Read-Host -prompt "What is the targetted UPN?"
[string]$internalMessage = Read-Host -prompt "What is the message for internal recipients?"
[string]$externalMessage = Read-Host -prompt "What is the message for external recipients?"

Set-MailboxAutoReplyConfiguration -identity $upn -AutoReplyState enabled -ExternalAudience all -InternalMessage $internalMessage -ExternalMessage $externalMessage
}

#creates a dynamic array and users a foreach loop which is dependent upon the amount of entered elements
function arrayWithValues{
#Will read the values entered as an array and then process. Think of it as dynamic array.
[string]$val = Read-Host -prompt "Please enter the values in the Array below:"
[string]$actualCommand = Read-Host -Prompt "Please enter the command you'd like to carry out"

foreach($val2 in $val){
Write-host "$val2"
$actualCommand
}
}

#checks for forwarding SMTP Addresses
function checkForwardingSMTP{
[string]$upn = Read-Host -prompt "What is the targetted UPN?"
get-mailbox -identity $upn | fl ForwardingAddress, ForwardingSMTPAddress
}
