#requires -modules ActiveDirectory
param ($Scope)

<#
.SYNOPSIS
  Provisioning script for QlikView and Citrix resources
.DESCRIPTION
  This script provides a tool to use AD groups to provision Qlikview Roles and Citrix Published Resources
.PARAMETER Scope
  QV: Qlikview
  CTX: Citrix
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Bart Jacobs - @Cloudsparkle
  Creation Date:  27/04/2021
  Purpose/Change: Qlikview/Citrix provisioning
 .EXAMPLE
  None
#>

# Function to read config.ini
Function Get-IniContent
{
    <#
    .Synopsis
        Gets the content of an INI file
    .Description
        Gets the content of an INI file and returns it as a hashtable
    .Notes
        Author        : Oliver Lipkau <oliver@lipkau.net>
        Blog        : http://oliver.lipkau.net/blog/
        Source        : https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version        : 1.0 - 2010/03/12 - Initial release
                      1.1 - 2014/12/11 - Typo (Thx SLDR)
                                         Typo (Thx Dave Stiff)
        #Requires -Version 2.0
    .Inputs
        System.String
    .Outputs
        System.Collections.Hashtable
    .Parameter FilePath
        Specifies the path to the input file.
    .Example
        $FileContent = Get-IniContent "C:\myinifile.ini"
        -----------
        Description
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent
    .Example
        $inifilepath | $FileContent = Get-IniContent
        -----------
        Description
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent
    .Example
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"
        C:\PS>$FileContent["Section"]["Key"]
        -----------
        Description
        Returns the key "Key" of the section "Section" from the C:\settings.ini file
    .Link
        Out-IniFile
    #>

    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)]
        [string]$FilePath
    )

    Begin
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"}

    Process
    {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"

        $ini = @{}
        switch -regex -file $FilePath
        {
            "^\[(.+)\]$" # Section
            {
                $section = $matches[1]
                $ini[$section] = @{}
                $CommentCount = 0
            }
            "^(;.*)$" # Comment
            {
                if (!($section))
                {
                    $section = "No-Section"
                    $ini[$section] = @{}
                }
                $value = $matches[1]
                $CommentCount = $CommentCount + 1
                $name = "Comment" + $CommentCount
                $ini[$section][$name] = $value
            }
            "(.+?)\s*=\s*(.*)" # Key
            {
                if (!($section))
                {
                    $section = "No-Section"
                    $ini[$section] = @{}
                }
                $name,$value = $matches[1..2]
                $ini[$section][$name] = $value
            }
        }
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"
        Return $ini
    }

    End
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"}
}

# Get ready for the GUI stuff
Add-Type -AssemblyName PresentationFramework

# Check for valid scope parameter
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

# Get the current running directory
$currentDir = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($currentDir -eq $PSHOME.TrimEnd('\'))
{
  $currentDir = $PSScriptRoot
}

# Read config.ini
$IniFilePath = $currentDir + "\config.ini"
$IniFileExists = Test-Path $IniFilePath
if ($IniFileExists -eq $true)
{
  $IniFile = Get-IniContent $IniFilePath
}
Else
{
  $msgBoxInput = [System.Windows.MessageBox]::Show("Config.ini not found.","Error","OK","Error")
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

# Get all AD users loaded in the beginning, so there is no delay further down
$ADusers = get-aduser -filter *  | select name, samaccountname | sort name

# If the scope is QV (for Qlikview), get everything ready
if ($scope -eq "QV")
{
  # Read the OU containing AD groups for Qlikview Roles
  $QVBaseOU = $IniFile["AD"]["QVBaseOU"]
  if ($QVBaseOU -eq $null)
  {
    $msgBoxInput = [System.Windows.MessageBox]::Show("OU for Qlikview Role Groups not found in config.ini.","Error","OK","Error")
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
    # Test if OU exists
    $QVBaseOUExists = Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$QVBaseOU'"
    if ($QVBaseOUExists -eq $null)
    {
      $msgBoxInput = [System.Windows.MessageBox]::Show("OU for Qlikview Role Groups not found in AD. Check config.ini.","Error","OK","Error")
      switch  ($msgBoxInput)
      {
        "OK"
        {
          Exit 1
        }
      }
    }
  }

  # Get the groups for this Scope
  $ADGroups = Get-ADGroup -Properties samaccountname, description -Filter '*' -SearchBase $QVBaseOU | sort description

  # Set all textlabels for this scope
  $Menu1 = "Add user to QlikView Role"
  $Menu2 = "Remove user from QlikView Role"
  $Menu3 = "List users for QlikView Role"
  $TitleSelectGroup = "Select the the QlikView Role:"
  $TitleRemoveUser = "Select the user to be removed from QlikView Role "
  $TitleListUser = "Users for QlikView Role "
  $TitleAddUser = "Select the user to be added to QlikView Role "
  $TextAddUser = " has been added to the selected Qlikview Role: "
  $TextRemoveUser = " has been removed from the selected Qlikview Role: "
}

if ($scope -eq "CTX")
{
  # Read the OU containing AD groups for Citrix Published resources
  $CTXBaseOU = $IniFile["AD"]["CTXBaseOU"]
  if ($CTXBaseOU -eq $null)
  {
    $msgBoxInput = [System.Windows.MessageBox]::Show("OU for Citrix Published Resources not found in config.ini.","Error","OK","Error")
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
    # Test if OU exists
    $CTXBaseOUExists = Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$CTXBaseOU'"
    if ($CTXBaseOUExists -eq $null)
    {
      $msgBoxInput = [System.Windows.MessageBox]::Show("OU for Citrix Published Resources not found in AD. Check config.ini.","Error","OK","Error")
      switch  ($msgBoxInput)
      {
        "OK"
        {
          Exit 1
        }
      }
    }
  }

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
    $TextAddUser = " has been added to the selected Citrix Published Resource: "
    $TextRemoveUser = " has been removed from the Citrix Published Resource: "
}

# Create the basic menu
$Menu = [ordered]@{

  1 = $Menu1
  2 = $Menu2
  3 = $Menu3
  }

# Get the chosen menu options and process it
$Result = $Menu | Out-GridView -Title 'Make a  selection' -OutputMode Single

Switch ($Result)
{
  {$Result.Name -eq 1}
  # Get through the steps to add a user to a QV/Citrix group
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
          exit 0
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
          exit 0
        }
      }
    }
  }

  {$Result.Name -eq 2}
  # Get through the steps to remove a user from a QV/Citrix group
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
        exit 0
      }
    }
    Remove-ADGroupMember -Identity $SelectedGroup.name -Members $SelectedUser.samaccountname -Confirm:$False
  }

  {$Result.Name -eq 3}
  # Get through the steps to list users for a QV/Citrix group
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
