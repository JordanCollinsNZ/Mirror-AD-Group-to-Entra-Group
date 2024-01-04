################################################################################
# Script Description
#
# This script can be used to add members of on-prem AD groups to Entra ID groups
# based off of the members email address. It will prompt the user to select an
# excel .xlsx file containing the groups with the below headers:
#
#  ADGroupName | EntraGroupName
# -------------+----------------
#  Group       | Group 
################################################################################
# Version History
#
# Version 1.0 - 12/09/23 - Jordan Collins - Initial creation
################################################################################
#
#Requires -Modules ActiveDirectory, Az.Resources, ImportExcel
#
################################################################################
# User set Variables
#
# Define AD domain name
    $ADDomainName = ""
################################################################################

# Define logging function
    function Log() {
    # Set function variables
        param (
            [Parameter(Mandatory = $True)] [String] $Outcome,
            [Parameter(Mandatory = $False)] [String] $IdP,
            [Parameter(Mandatory = $False)] [String] $Group,
            [Parameter(Mandatory = $True)] [String] $LogMessage
        )

    # Create log entry object
        $LogEntry = New-Object PSObject -Property @{
            Outcome = $Outcome
            IdP = $IdP
            Group = $Group
            LogMessage = $LogMessage
        }

    # Write log messages to host
        Write-Host "$Outcome - $LogMessage"

    # Write log messages to the log file
        $LogEntry | Export-Excel -Path $LogFilePath -Append
    }

# Define Log variables and Log Path
    $DateTime = Get-Date -Format "dd.MM.yy HH.mm"
    $LogFilePath = ".\Mirror Log $DateTime.xlsx"

# Import required modules
    Import-Module Az.Resources
    Import-Module ActiveDirectory
    Import-Module ImportExcel

# Define AD server based off domain
    try {
        $ADServer = (Get-ADDomain -Identity "$ADDomainName").InfrastructureMaster
    } catch {
        throw "Cannot get infrastructure master for Active Directory domain $ADDomainName."
    }

# Create an OpenFileDialog box
    [Void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [Void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title            = "Select XLSX File"
        InitialDirectory = [Environment]::GetFolderPath("Desktop")
        Filter           = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    }

# Show the dialog and wait for user selection
    Write-Host "Select the Excel .xlsx file in the new popup."
    $DialogResult = $OpenFileDialog.ShowDialog([System.Windows.Forms.Form]::WindowState)
    if ($DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "File selection cancelled."
    }

# Get the selected file path
    $ExcelFilePath = $OpenFileDialog.FileName

# Read group names from the selected XLSX file
    $GroupData = Import-Excel -Path $ExcelFilePath

# Connect to Entra ID. I am using Device auth as I came across weird bugs where Connect-AzAccount would hang after logging in
    try {
        Start-Process https://login.microsoftonline.com/common/oauth2/deviceauth
        Connect-AzAccount -UseDeviceAuthentication
    } catch {
         throw "Failed to authenticate with Entra ID. Please try again."
    }

# Reset Progress Counter
    $ProgressCounter0 = 0   

# Loop through each row in the XLSX file
    foreach ($Row in $GroupData) {
        # Read column headers to variables
            $EntraGroupDisplayName = $Row.EntraGroupName
            $ADGroupName = $Row.ADGroupName

        # Set progress counter status
            $ProgressPercent0 = [Math]::Round((($ProgressCounter0/($GroupData | Measure-Object).count) * 100), 0)
			Write-Progress -ID 0 -Activity "Groups Progress" -Status "$ProgressPercent0% Complete" -CurrentOperation $ADGroupName -PercentComplete $ProgressPercent0

        # Get the Entra group based on its display name
            $EntraGroup = Get-AzADGroup -WarningAction Ignore -DisplayName $EntraGroupDisplayName

        # Check if the Entra group exists
            if ($null -eq $EntraGroup) {
                Log -Outcome "Skipped" -IdP "Entra" -Group "$($EntraGroup.DisplayName)" -LogMessage "Entra Group $($EntraGroup.DisplayName) specified in spreadsheet not found in Entra ID"
                continue
            }

        # Get existing Entra Group members
            $EntraGroupMembers = (Get-AzADGroupMember -WarningAction Ignore -GroupObjectId $EntraGroup.Id | Select-Object -ExpandProperty Id)

        # Get members of the AD group
            $ADGroupMembers = (Get-ADGroupMember -Identity $ADGroupName -Recursive -Server $ADServer | Get-ADUser -Properties EmailAddress | Select-Object UserPrincipalName, ObjectClass, EmailAddress)

        # Reset Progress counter
			$ProgressCounter1 = 0

        # Loop through the members of the AD group
            foreach ($ADMember in $ADGroupMembers) {
                # Set progress counter status
					$ProgressPercent1 = [Math]::Round((($ProgressCounter1/$ADGroupMembers.count) * 100), 0)
					Write-Progress -Id 1 -ParentId 0 -Activity "Group Progress" -Status "$ProgressPercent1% Complete" -CurrentOperation $ADMember.UserPrincipalName -PercentComplete $ProgressPercent1

                # Check if the member is a user
                    if ($ADMember.ObjectClass -eq "User") {
                        # Not really needed, but change variable to user to stay consistant now that the member is confirmed as a user
                            $ADUser = $ADMember
                        # Check if the user has an email set
                            if ($null -ne $ADUser.EmailAddress) {
                                # Set variable for users email in Entra ID    
                                    $EntraUser = (Get-AzAdUser -WarningAction Ignore -Mail $ADUser.EmailAddress)
                                # Check if user found in Entra ID
                                    if ($null -ne $EntraUser) {
                                        # Check if user is already a member of Entra group
                                            if ($EntraGroupMembers -notcontains $EntraUser.Id) {
                                                # Add the user to the Entra group
                                                    Add-AzADGroupMember -WarningAction Ignore -TargetGroupObjectId $EntraGroup.Id -MemberObjectId $EntraUser.Id
                                                # Add success to log
                                                    Log -Outcome "Success" -Group "$($EntraGroup.DisplayName)" -LogMessage "Added $($EntraUser.Mail) to $($EntraGroup.DisplayName)."
                                            } else {
                                                # Skip adding user to group and log
                                                    Log -Outcome "Skipped" -Group "$($EntraGroup.DisplayName)" -LogMessage "$($EntraUser.Mail) is already a member of $($EntraGroup.DisplayName)."
                                            }
                                    } else {
                                        # Cannot find user with the specified email in Entra. Skip and log
                                            Log -Outcome "Not Found" -Group "$($EntraGroup.DisplayName)" -IdP "Entra" -LogMessage "$($ADUser.EmailAddress) not found in Entra ID."
                                    }
                            } else {
                                # Member user has no email set in Active Directory. Skip and log.
                                    Log -Outcome "Not Set" -Group "$ADGroupName" -IdP "Active Directory" -LogMessage "$($ADUser.UserPrincipalName) has no email set."
                            }
                    } else {
                        # Member of group is not a user. Skip and log.
                            Log -Outcome "Skipped" -IdP "Active Directory" -Group "$ADGroupName" -LogMessage "Skipping non-user member $($ADMember.UserPrincipalName)"
                    }
                # Add to Group Progress counter
					$ProgressCounter1++
            }
        #Add to Groups Progress counter
			$ProgressCounter0++
    }

# Disconnect from Entra ID
    Disconnect-AzAccount | Out-Null

Write-Host "Script execution complete. Log file saved at $LogFilePath."
