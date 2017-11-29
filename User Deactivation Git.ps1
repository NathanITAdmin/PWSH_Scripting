#This script will ask for the user name of the user that you wish to disable.
#The script will then disable the account, remove it from all groups and changes the password.
#The script then prompts for and uses admin credentials to log into O365.
#It Sets the away message to tell them to contact the user of your choice and sets up forwarding if needed. 
#If forwarding is needeed it will convert the mailbox into a shared mailbox and remove the license from the user. If it it not needed then it just removes the license.
#
# Author: Nathan Medeiros
# Version: 1.0
# Last Change Date: 28FEB2017 


#Get account info for the account that is to be disabled
$SamName = Read-Host "Please enter the username of the user that is to be disabled"
$ExitUser = Get-ADUser $SamName -Properties *
$UPN = Get-ADUser $SamName -Properties UserPrincipalName | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName 

function PrintError($strMsg)
{
    Write-Host $strMsg -foregroundcolor "red"
}

function PrintSuccess($strMsg)
{
    Write-Host $strMsg -foregroundcolor "green"
}

# Cleans up and prints an error message
function CleanupAndFail($strMsg)
{
    if ($strMsg)
    {
        PrintError($strMsg);
    }
    Cleanup
    exit 1
}
function ExitIfError($strMsg)
{
    if ($Error)
    {
        CleanupAndFail($strMsg);
    }
}
Function InPutNeeded($strMsg)
{
    $AskQuest = Read-Host $strMsg
    if ($AskQuest -eq "Y")
    {
        $YesNo = $true
    }
    elseif ($AskQuest -eq "N")
    {
        $YesNo = $false
    }
    else
    {
        $YesNo = InPutNeeded($strMsg)
    }
    Return $YesNo
}
$Active = InPutNeeded("Will the users email need to be accessed after they leave? Please select 'Y' or 'N'")
if ($Active)
{
    Get-ADUser $SamName | Move-ADObject -TargetPath "OU=OUNAME,OU=OUNAME,DC=Company,DC=com"
}
else 
{
    Get-ADUser $SamName | Move-ADObject -TargetPath "OU=OUNAME,DC=Company,DC=com"
}

#Resets user password and turns off change password at next logon
Get-ADUser $SamName | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "PassWord" -Force)

Set-ADUser -Identity $SamName -ChangePasswordAtLogon $false

#Remove user from all groups
Get-ADUser $SamName -Properties memberof | Select-Object memberof -ExpandProperty memberof | Remove-ADGroupMember -Members $SamName -confirm:$false

#Disable user account in AD
Disable-ADAccount -Identity $SamName

## Sign in to remote powershell for exchange and lync online ##
Write-Host "`------------------  Establishing connection  -----------------." -foregroundcolor "magenta"
$credAdmin=Get-Credential -Message "Enter credentials of a O365 admin"
if (!$credadmin)
{
    CleanupAndFail("Valid admin credentials are required to remove the account.");
}
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $credAdmin -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
Import-PSSession $Session
#Asks if emails will need to be forwarded and to whom
if ($Active)
{
    Set-Mailbox -Identity $SamName -Type Shared

    $Forward = Read-Host "Please enter the email of the user that emails will be forwarded to. Example: John.Doe@Company.com"

    Set-Mailbox -Identity $SamName -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $Forward
}

#Asks if an away message will be needed ands sets it if needed
$Need = InPutNeeded("Will the users email need an away message directing them to a new contact? (Y)es or (N)o? Please select 'Y' or 'N'")
if($Need)
{
    $Destination = Read-Host "Please enter the user name of the user you would like the away message to direct people to."
}

if($Destination)
{
    $Relay = Get-ADUser $Destination -Properties *
    $firstout = $ExitUser.givenname
    $FLFWD = $Relay.name
    $EmailFWD = $Relay.mail

    Set-MailboxAutoReplyConfiguration -Identity $SamName -AutoReplyState Enabled `
    -InternalMessage "$firstout is no longer with Company. If you have any questions or concerns please contact $FLFWD at $EmailFWD"
}

#Connects to MSOL to remove license from user
Connect-MsolService -Credential $credAdmin

Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "O365License"




