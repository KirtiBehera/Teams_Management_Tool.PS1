<#
=============================================================================================
Name:           Teams Management Tool
Description:    This script handles Teams-related tasks and generates reports. 
Scripted by:    Kirti R Behera
============================================================================================
#>
param (
    [string]$Title = "Select option for MS Teams Management",
    [int]$Action
)
 #Install-Module -Name MicrosoftTeams -Confirm:$false
Connect-MicrosoftTeams

[boolean]$Delay = $false
Do {
    if ($Action -eq "") {
        if ($Delay -eq $true) {
            Start-Sleep -Seconds 2
        }
        $Delay = $true

        Write-Host ""
        Write-Host "`nMicrosoft Teams Create" -ForegroundColor Yellow
        Write-Host "    1. Creating New Teams" -ForegroundColor Cyan
        Write-Host "    2. Creating New Teams Channel" -ForegroundColor Cyan
        Write-Host "`nMicrosoft Teams Modification" -ForegroundColor Yellow
        Write-Host "    3. Add Members to Teams" -ForegroundColor Cyan
        Write-Host "    4. Add Members to Teams Channel" -ForegroundColor Cyan
        Write-Host "    5. Change Display Name" -ForegroundColor Cyan
        Write-Host "    6. Set Teams Picture" -ForegroundColor Cyan
        Write-Host "`nMicrosoft Teams Deletion" -ForegroundColor Yellow
        Write-Host "    7. Remove Teams or Teams Channel" -ForegroundColor Cyan
        Write-Host "    8. Remove Teams User or Teams Channel User" -ForegroundColor Cyan
        Write-Host "`nMicrosoft Teams Restore" -ForegroundColor Yellow
        Write-Host "    9. Restore Teams Group" -ForegroundColor Cyan
        Write-Host "`nBulk process" -ForegroundColor Yellow
        Write-Host "    10. Add Bulk Members to Teams or Teams channel" -ForegroundColor Cyan
        Write-Host "    11. Remove Bulk Members from Teams or Teams channel" -ForegroundColor Cyan
        Write-Host "`nTeams Channel Reporting" -ForegroundColor Yellow
        Write-Host "    12. Members and Owners Report of Single TeamsGroup" -ForegroundColor Cyan
        Write-Host "    13. Members and Owners Report of Single Channel" -ForegroundColor Cyan
        Write-Host "`nExit from Script" -ForegroundColor Yellow
        Write-Host "    0. Exit" -ForegroundColor Cyan
        Write-Host ""

        $i = Read-Host 'Please choose the action to continue' 
    } else {
        $i = $Action
    }
 
 Switch ($i) {
  1 {
     $display = Read-Host "Enter Display Name" 
     $description = Read-Host "Enter Description" 
     New-Team -DisplayName $display -Description $description -Visibility Public
     Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green   
    }
      
  2 {
     $group = Read-Host "Enter Teams Group Name"
     $display = Read-Host "Enter Display Name" 
     $groupid = (Get-Team -DisplayName $group).GroupId 
     New-TeamChannel -GroupId $groupid -DisplayName $display -MembershipType Private
     Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green 
    }

  3 {
     $group = Read-Host "Enter Teams Group Name"
     $member = Read-Host "Enter members email address" 
     $role = Read-Host "Enter members Role (Owner/Member)"
     $groupid = (Get-Team -DisplayName $group).GroupId 
     Add-TeamUser -GroupId $groupid -User $member -Role $role
     Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
     }  
 4 {
    $group = Read-Host "Enter Group Name"
    $channel = Read-Host "Enter Channel Name"
    $member = Read-Host "Enter members email address" 
    #$role = Read-Host "Enter members Role (Owner/Member)"
    $groupid = (Get-Team -DisplayName $group).GroupId 
    Add-TeamChannelUser -GroupId $groupid -DisplayName $channel -User $member 
    Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
     }

  5 {
    $group = Read-Host "Enter Group Name"
    $newgroup = Read-Host "Enter New Group Name"
    $groupid = (Get-Team -DisplayName $group).GroupId
    $disp = Read-host "Select 1 for "Teams" and 2 for "Teams Channel" " 
    if ($disp -eq 1) { Set-Team -GroupId $groupid -DisplayName $newgroup -Visibility Public }
    if ($disp -eq 2) { Set-TeamChannel -GroupId $groupid -CurrentDisplayName $group -NewDisplayName $newgroup }
    Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
      }   

  6 {
     $group = Read-Host "Enter Group Name"
     $imagelocation = Read-Host "Enter image path (ex:c:\Image\TeamsPicture.png)"
     $groupid = (Get-Team -DisplayName $group).GroupId 
     Set-TeamPicture -GroupId $groupid -ImagePath $imagelocation
     Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green 
    }
       
   7 {
    $group = Read-Host "Enter Group Name"
    $groupid = (Get-Team -DisplayName $group).GroupId 
    $Teams = Read-host "Select 1 for "Remove Teams" and 2 for "Remove Teams Channel" "
    if ($Teams -eq 1) { Write-Host "Removing Teams group "
    Remove-Team -GroupId $groupid }
    if ($Teams -eq 2) { Write-Host "Removing Teams Channel " 
    $channel = Read-Host "Enter channel Name"
    Remove-TeamChannel -GroupId $groupid -DisplayName $channel }
    Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
    }  

   8 {
    $group = Read-Host "Enter Group Name"
    $member = Read-Host "Enter members email address"
    $groupid = (Get-Team -DisplayName $group).GroupId 
    $user = Read-host "Select 1 for "Remove Teams User" and 2 for "Remove Teams Channel User" " 
    if ($user -eq 1) { Write-Host "Removing Teams user " 
    Remove-TeamUser -GroupId $groupid -User $member }
    if ($user -eq 2) { Write-Host "Removing Teams Channel user" 
    $channel = Read-Host "Enter Channel Name"
    Remove-TeamChannelUser -GroupId $groupid -DisplayName $channel -User $member }
    Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
      }

  9 {
  Install-Module -Name ExchangeOnlineManagement -Force
  Connect-ExchangeOnline
  $group = Read-Host "Enter Group Name"
    Undo-SoftDeletedUnifiedGroup -SoftDeletedObject $group
    Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
   }
0 {
   Write-Host "Exiting script." -ForegroundColor Green
            exit 
    }
    Default {
            Write-Host "Invalid selection. Please choose a valid action." -ForegroundColor Red
       }

10 {
$CSVPath = ".\TeamsUsersTemplate.csv"
$option = Read-host "Select 1 for Teams Group and 2 for Teams Channel "
if ($option -eq 1) { 
    Write-Host "Adding Bulk Member to Teams Group"
    $TeamDisplayName = Read-Host "Enter Group Name"
    Try { 
        $TeamID = Get-Team -DisplayName $TeamDisplayName | Select -ExpandProperty GroupID
        $TeamUsers = Import-Csv -Path $CSVPath
        $TeamUsers | ForEach-Object {
            Try {
                Add-TeamUser -GroupId $TeamID -User $_.Email -Role $_.Role
                Write-host "Added User:"$_.Email -f Green
            }
            Catch {
                Write-host -f Red "Error Adding User to the Team Group:" $_.Exception.Message
            }
        }
    }
    Catch {
        Write-host -f Red "Error:" $_.Exception.Message
    }
}
if ($option -eq 2) { 
    Write-Host "Adding Bulk Member to Teams Channel"
    $TeamDisplayName = Read-Host "Enter Group Name"
    $channel = Read-Host "Enter Channel Name"
    Try { 
        $TeamID = Get-Team -DisplayName $TeamDisplayName | Select -ExpandProperty GroupID
        $TeamUsers = Import-Csv -Path $CSVPath
        $TeamUsers | ForEach-Object {
            Try {
                Add-TeamChannelUser -GroupId $TeamID -DisplayName $channel -User $_.Email 
                Write-host "Added User to Teams Channel :"$_.Email -f Green
            }
            Catch {
                Write-host -f Red "Error Adding User to the Team Channel:" $_.Exception.Message
            }
        }
    }
    Catch {
        Write-host -f Red "Error:" $_.Exception.Message
    }
}
Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
  }  

 11 {
 $CSVPath = ".\TeamsUsersTemplate.csv"
$option = Read-host "Select 1 for Teams Group and 2 for Teams Channel "
if ($option -eq 1) { 
    Write-Host "Removing Bulk Member from Teams Group"
    $TeamDisplayName = Read-Host "Enter Group Name"
    Try { 
        $TeamID = Get-Team -DisplayName $TeamDisplayName | Select -ExpandProperty GroupID
        $TeamUsers = Import-Csv -Path $CSVPath
        $TeamUsers | ForEach-Object {
            Try {
                Remove-TeamUser -GroupId $TeamID -User $_.Email
                Write-host "Removed User:"$_.Email -f Green
            }
            Catch {
                Write-host -f Red "Error Removing User From Team Group:" $_.Exception.Message
            }
        }
    }
    Catch {
        Write-host -f Red "Error:" $_.Exception.Message
    }
}
if ($option -eq 2) { 
    Write-Host "Removing Bulk Member from Teams Channel"
    $TeamDisplayName = Read-Host "Enter Group Name"
    $channel = Read-Host "Enter Channel Name"
    Try { 
        $TeamID = Get-Team -DisplayName $TeamDisplayName | Select -ExpandProperty GroupID
        $TeamUsers = Import-Csv -Path $CSVPath
        $TeamUsers | ForEach-Object {
            Try {
            Remove-TeamChannelUser -GroupId $TeamID -DisplayName $channel -User $_.Email
                Write-host "Removed User from Teams Channel :"$_.Email -f Green
            }
            Catch {
                Write-host -f Red "Error Removing User from Team Channel:" $_.Exception.Message
            }
        }
    }
    Catch {
        Write-host -f Red "Error:" $_.Exception.Message
    }
}
Write-Host "Kindly allow 30Mins to 60Mins for replication" -ForegroundColor Green
 }  
 
 12 {
    $Result = ""  
    $Results = @()     
      $TeamName=Read-Host Enter Teams name "(Case Sensitive)"
      $Email = Read-Host "Enter Email address for Reports" 
      Write-Host Exporting Channels report...
      #$Count=0
      $Team = Get-Team -DisplayName $TeamName
if ($Team -eq $null) {
    Write-Host "Team '$TeamName' not found."
    exit
}
$GroupId = $Team.GroupId

      #$GroupId=(Get-Team -DisplayName $TeamName).GroupId
      $Path=".\Channels available in $TeamName team $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      $AllResults = @()

      $Channels = Get-TeamChannel -GroupId $GroupId

      $htmlContent = @"
<style>
    table {border-collapse: collapse; width: 100%;}
    th, td {border: 1px solid black; padding: 8px; text-align: left;}
</style>
"@

foreach ($Channel in $Channels) {
    $ChannelName = $Channel.DisplayName
    Write-Progress -Activity "Processed channel count: $($AllResults.Count)" -Status "Currently Processing Channel: $ChannelName"

    # Retrieve channel details
    $MembershipType = $Channel.MembershipType
    $Description = $Channel.Description

    # Get all users in the channel
    $ChannelUsers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName

    # Count of channel members
    $ChannelMemberCount = $ChannelUsers.Count

    # Prepare results for each user in the channel
    foreach ($User in $ChannelUsers) {
        $UserResult = [PSCustomObject]@{
            'Teams Name' = $TeamName
            'Channel Name' = $ChannelName
            'Membership Type' = $MembershipType
            'Description' = $Description
            'Total Members Count' = $ChannelMemberCount
            'Member Name' = $User.Name
            'Member Role' = $User.Role
            'Member Email' = $User.user
        }     
       $AllResults += $UserResult
        }
                

}

# Export to CSV
        $AllResults | Export-Csv -Path $Path -NoTypeInformation -Append

        # Export to HTML (optional)
        $htmlTable = $AllResults| ConvertTo-Html -Head $htmlContent -Property 'Teams Name', 'Channel Name', 'Membership Type', 'Description', 'Total Members Count', 'Member Name', 'Member Role', 'Member Email'
        
        Send-MailMessage -To $Email  -CC "kirtiranjan@conteso.com" -From "Teams.Reports@conteso.com" -Subject "Teams Group report: $TeamName" -SmtpServer "192.168.1.1" -Body "<h1>Teams Group: $TeamName Report</h1><br><br>$htmlTable" -BodyAsHtml

      Write-Progress -Activity "`n     Processed channel count: $Count "`n"  Currently Processing Channel: $ChannelName" -Completed
      if((Test-Path -Path $Path) -eq "True") 
      {
       Write-Host `nReport available in $Path -ForegroundColor Green
      }
     }  

  
    13 {
    $Result = ""  
    $Results = @() 
    $TeamName = Read-Host "Enter Teams name in which Private Channel resides (Case sensitive)"
    $ChannelName = Read-Host "Enter Private Channel name"
    $Email = Read-Host "Enter Email address for Reports" 
    
    $GroupId = (Get-Team -DisplayName $TeamName).GroupId 
    Write-Host "Exporting $ChannelName's Members and Owners report..." -ForegroundColor Green
    $Path = ".\MembersOf_$ChannelName_$((Get-Date -format yyyy-MMM-dd-ddd_hh-mm_tt)).csv"

    $htmlContent = @"
<style>
    table {border-collapse: collapse; width: 100%;}
    th, td {border: 1px solid black; padding: 8px; text-align: left;}
</style>
"@
    
    $Report = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach {
     $Name = $_.Name
     $UPN = $_.User
     $Role = $_.Role
      [PSCustomObject]@{
            'Teams Name' = $TeamName
            'Private Channel Name' = $ChannelName
            'Channel Members Email Address' = $UPN
            'User Display Name' = $Name
            'Role' = $Role
        }
     }
        
   $Report | Export-Csv -Path $Path -NoTypeInformation
    $htmlTable = $Report | ConvertTo-Html -Head $htmlContent -Property 'Teams Name', 'Private Channel Name', 'Channel Members Email Address', 'User Display Name', 'Role'

    Send-MailMessage -To $Email  -CC "kirtiranjan.behera@conteso.com" -From "Teams.Reports@conteso.com" -Subject "Teams Group Channel report : $ChannelName" -SmtpServer "192.168.1.00" -Body "<h1>Teams Group: $ChannelName Report</h1><br><br>$htmlTable" -BodyAsHtml

    if (Test-Path -Path $Path) {
        Write-Host "`nReport available in $Path" -ForegroundColor Green}
   }
 }
  } While ($true) 

 
