#requires -modules ActiveDirectory
param ($Scope)

<#
.SYNOPSIS
  Automatically launch a Citrix published resource
.DESCRIPTION
  This script provides a GUI to quickly clean an existing AD group from disabled user accounts, optionally recursive
.PARAMETER PublishedApp
  Name of the Citrix published resource
.INPUTS
  Name of Citrix published resource
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Bart Jacobs - @Cloudsparkle
  Creation Date:  16/02/2021
  Purpose/Change: Automatically launch a Citrix published resource
 .EXAMPLE
  None
#>

Add-Type -AssemblyName PresentationFramework

if (($scope -ne "QV") -and ($scope -ne "CTX"))
{
    $msgBoxInput =  [System.Windows.MessageBox]::Show("Scope parameter missing or invalid.","Error","OK","Error")

    switch  ($msgBoxInput) {

    "OK" {

        Exit 1 

        }
        }
    
}

#
$SelectedGroup=""
$SelectedUser=""
$ADGroups=""

$QVBaseOU = "OU=Qlikview,OU=Security Groups,OU=Nitto Europe NV,DC=nittoeurope,DC=com"
$CTXBaseOU = "OU=CTX,OU=Groups,OU=Fujitsu,DC=nittoeurope,DC=com"
$ADusers = get-aduser -filter *  | select name, samaccountname | sort name

$SelectedDomain = Get-ADDomain
#Find the right AD Domain Controller
$dc = Get-ADDomainController -DomainName $SelectedDomain.Forest -Discover -NextClosestSite

if ($scope -eq "QV")
{
   
    $ADGroups = Get-ADGroup -Properties samaccountname, info, description -Filter '*' -SearchBase $QVBaseOU 
    
    $Menu1 = "Add user to role"
    $Menu2 = "Remove user from role"
    $Menu3 = "List users for role"
    $TitleSelectGroup = "Select the the role"
    $TitleRemoveUser = "Select the user to be removed from role"
    $TitleAddUser = "Select the user to be added to role"
}

if ($scope -eq "CTX")
{
   $ADGroups = Get-ADGroup -Properties samaccountname, info, description -Filter '*' -SearchBase $CTXBaseOU 
    
    $Menu1 = "Add user to Citrix Published Resource"
    $Menu2 = "Remove user from Citrix Published Resource"
    $Menu3 = "List users for Citrix Published Resource"
    $TitleSelectGroup = "Select the the Citrix Published Resource"
    $TitleRemoveUser = "Select the user to be removed from Citrix Published Resource"
    $TitleAddUser = "Select the user to be added to Citrix Published Resource"
}

$Menu = [ordered]@{

  1 = $Menu1
  2 = $Menu2
  3 = $Menu3
  }

$Result = $Menu | Out-GridView -Title 'Make a  selection' -OutputMode Single

  Switch ($Result)  {

  {$Result.Name -eq 1} 
    {
        $SelectedGroup = $ADGroups| select description, name | Out-GridView -Title $TitleSelectGroup -OutputMode Single
        if ($SelectedGroup -eq $null)
        {
            exit 0
        }
        $SelectedUser = $ADusers | Out-GridView -Title $TitleAddUser -OutputMode Single
        if ($SelectedUser -eq $null)
        {
            exit 0
        }
        $ADGroupMembers = Get-ADGroupMember -Identity $SelectedGroup.name | Select-Object -ExpandProperty samAccountName
        if ($SelectedUser.SamAccountName -in $ADGroupMembers)
        {
            $msgBoxInput =  [System.Windows.MessageBox]::Show("Selected user is already a member. Exiting now.","Warning","OK","Warning")

            switch  ($msgBoxInput) {

            "OK" {

                    Exit 1 

                    }
             }
            
            
        }
        Else
        {
            Add-ADGroupMember -Identity $SelectedGroup.name -Members $Selecteduser.SamAccountName
        }
    }

  {$Result.Name -eq 2} 
  {
    $SelectedGroup = $ADGroups| select description, name | Out-GridView -Title $TitleSelectGroup -OutputMode Single
    if ($SelectedGroup -eq $null)
        {
            exit 0
        }
    $SelectedUser = Get-ADGroupMember -Identity $SelectedGroup.name | select Name, SamAccountName | Out-GridView -Title $TitleRemoveUser -OutputMode Single
    if ($SelectedUser -eq $null)
        {
            exit 0
        }
    Remove-ADGroupMember -Identity $SelectedGroup.name -Members $SelectedUser.samaccountname -Confirm:$False
    }

  
} 