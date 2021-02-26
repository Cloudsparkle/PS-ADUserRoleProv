#requires -modules ActiveDirectory
param ($Scope)

<#
.SYNOPSIS
  Provisioning script for QlikView and Citrix resources
.DESCRIPTION
  This script provides a tool to use AD groups to provision Qlikview Roles and Citrix published resources
.PARAMETER Scope
  Qlikview or Citrix
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Bart Jacobs - @Cloudsparkle
  Creation Date:  26/02/2021
  Purpose/Change: Qlikview/Citrix provisioning
 .EXAMPLE
  None
#>

# Get ready for the GUI stuff
Add-Type -AssemblyName PresentationFramework

if (($scope -ne "QV") -and ($scope -ne "CTX"))
{
  $msgBoxInput =  [System.Windows.MessageBox]::Show("Scope parameter missing or invalid.","Error","OK","Error")
  switch  ($msgBoxInput)
  {
    "OK"
    {
      Exit 1
    }
  }
}

# Initialize variables
$SelectedGroup=""
$SelectedUser=""
$ADGroups=""

# Set the base OU to search for groups
$QVBaseOU = "OU=Qlikview,OU=Security Groups,OU=Nitto Europe NV,DC=nittoeurope,DC=com"
$CTXBaseOU = "OU=CTX,OU=Groups,OU=Fujitsu,DC=nittoeurope,DC=com"
# Get all AD users loaded in the beginning, so there is no delay further down
$ADusers = get-aduser -filter *  | select name, samaccountname | sort name

if ($scope -eq "QV")
{
  # Get the groups for this Scope
  $ADGroups = Get-ADGroup -Properties samaccountname, description -Filter '*' -SearchBase $QVBaseOU | sort description

  # Set all textlabels for this scope
  $Menu1 = "Add user to QlikView role"
  $Menu2 = "Remove user from QlikView role"
  $Menu3 = "List users for QlikView role"
  $TitleSelectGroup = "Select the the QlikView role:"
  $TitleRemoveUser = "Select the user to be removed from QlikView role "
  $TitleListUser = "Users for QlikView role "
  $TitleAddUser = "Select the user to be added to QlikView role "
  $TextAddUser = " has been added to the selected role "
  $TextRemoveUser = " has been removed from the selected role "
}

if ($scope -eq "CTX")
{
  # Get the groups for this Scope
  $ADGroups = Get-ADGroup -Properties samaccountname, description -Filter '*' -SearchBase $CTXBaseOU  | sort description

  # Set all textlabels for this scope
  $Menu1 = "Add user to Citrix Published Resource"
  $Menu2 = "Remove user from Citrix Published Resource"
  $Menu3 = "List users for Citrix Published Resource"
  $TitleSelectGroup = "Select the the Citrix Published Resource:"
  $TitleRemoveUser = "Select the user to be removed from Citrix Published Resource "
  $TitleAddUser = "Select the user to be added to Citrix Published Resource "
  $TitleListUser = "Users for Citrix Published Resource "
  $TextAddUser = " has been added to the selected Citrix Published Resource "
  $TextRemoveUser = " has been removed from the Citrix Published Resource "
}

$Menu = [ordered]@{

  1 = $Menu1
  2 = $Menu2
  3 = $Menu3
  }

$Result = $Menu | Out-GridView -Title 'Make a  selection' -OutputMode Single

Switch ($Result)
{
  {$Result.Name -eq 1}
  {
    $SelectedGroup = $ADGroups| select description, name | Out-GridView -Title $TitleSelectGroup -OutputMode Single
    if ($SelectedGroup -eq $null)
    {
      exit 0
    }
    $SelectedUser = $ADusers | Out-GridView -Title ($TitleAddUser + $SelectedGroup.description) -OutputMode Single
    if ($SelectedUser -eq $null)
    {
      exit 0
    }
    $ADGroupMembers = Get-ADGroupMember -Identity $SelectedGroup.name | Select-Object -ExpandProperty samAccountName
    if ($SelectedUser.SamAccountName -in $ADGroupMembers)
    {
      $msgBoxInput =  [System.Windows.MessageBox]::Show("Selected user is already a member. Exiting now.","Warning","OK","Warning")
      switch  ($msgBoxInput)
      {
        "OK"
        {
          Exit 1
        }
      }
    }
    Else
    {
      Add-ADGroupMember -Identity $SelectedGroup.name -Members $Selecteduser.SamAccountName
      $msgBoxText = $SelectedUser.name + $TextAddUser + $SelectedGroup.description
      $msgBoxInput =  [System.Windows.MessageBox]::Show($msgBoxText,"User added","OK","Information")
      switch  ($msgBoxInput)
      {
        "OK"
        {
          Exit 1
        }
      }
    }
  }

  {$Result.Name -eq 2}
  {
    $SelectedGroup = $ADGroups| select description, name | Out-GridView -Title $TitleSelectGroup -OutputMode Single
    if ($SelectedGroup -eq $null)
    {
      exit 0
    }
    $SelectedUser = Get-ADGroupMember -Identity $SelectedGroup.name | select Name, SamAccountName | Sort Name | Out-GridView -Title ($TitleRemoveUser + $SelectedGroup.description) -OutputMode Single
    if ($SelectedUser -eq $null)
    {
      exit 0
    }
    $msgBoxText = $SelectedUser.name + $TextRemoveUser + $SelectedGroup.description
    $msgBoxInput =  [System.Windows.MessageBox]::Show($msgBoxText,"User removed","OK","Information")
    switch  ($msgBoxInput)
    {
      "OK"
      {
        Exit 1
      }
    }
    Remove-ADGroupMember -Identity $SelectedGroup.name -Members $SelectedUser.samaccountname -Confirm:$False
  }

  {$Result.Name -eq 3}
  {
    $SelectedGroup = $ADGroups| select description, name | Out-GridView -Title $TitleSelectGroup -OutputMode Single
    if ($SelectedGroup -eq $null)
    {
      exit 0
    }
    $SelectedUser = Get-ADGroupMember -Identity $SelectedGroup.name | select Name, SamAccountName | sort Name | Out-GridView -Title ($TitleListUser + $SelectedGroup.description) -OutputMode Single
    exit 0
  }

}
