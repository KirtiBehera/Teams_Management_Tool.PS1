# TeamsManagementTool.ps1

## Overview

**TeamsManagementTool.ps1** is a PowerShell script that simplifies Microsoft Teams management. It allows users to automate various tasks related to creating, managing, modifying, and deleting Teams groups, channels, and members. This script reduces administrative effort and enhances the organization's workflow by streamlining Microsoft Teams operations. It is a step toward automation and AI-driven systems.

## Features

- **Create Teams Groups and Channels**
  - Create new Teams groups and channels with ease.
- **Manage Teams Groups and Channels**
  - Modify existing groups, channels, and team members.
- **Modify Teams Structure**
  - Update group and channel settings, add/remove members, and modify group images.
- **Delete Teams Groups, Channels, and Members**
  - Safely delete groups, channels, and remove members.
- **Restore Teams Groups**
  - Restore deleted Teams groups as needed.
- **Automation and AI-Driven**
  - This tool brings automation to Microsoft Teams management, providing AI-driven optimization for handling administrative tasks.

## Prerequisites

Before running the script, ensure you have the following installed:

- **PowerShell** (v5.1 or later)
- **Microsoft Teams PowerShell Module**  
  Install it by running the following command in PowerShell:
  ```powershell
  Install-Module -Name MicrosoftTeams -Force -AllowClobber
  ```
- **Admin Permissions**  
  Ensure you have administrative rights to perform Teams management tasks.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/KirtiBehera/Teams_Management_Tool.PS1.git
   ```

2. Change directory to the repository:
   ```bash
   cd TeamsManagementTool
   ```

3. Ensure the necessary modules are installed:
   ```powershell
   Install-Module -Name MicrosoftTeams -Force -AllowClobber
   ```

## Usage

Run the PowerShell script with elevated privileges to manage Microsoft Teams.

### Example Commands:

1. **Create a Teams Group:**
   ```powershell
   .\TeamsManagementTool.ps1 -CreateGroup -GroupName "Your Group Name"
   ```

2. **Create a Channel within a Group:**
   ```powershell
   .\TeamsManagementTool.ps1 -CreateChannel -GroupName "Your Group Name" -ChannelName "Channel Name"
   ```

3. **Modify a Group:**
   ```powershell
   .\TeamsManagementTool.ps1 -ModifyGroup -GroupName "Your Group Name" -NewGroupName "Updated Group Name"
   ```

4. **Add a Member to a Group:**
   ```powershell
   .\TeamsManagementTool.ps1 -AddMember -GroupName "Your Group Name" -MemberEmail "user@example.com"
   ```

5. **Remove a Channel:**
   ```powershell
   .\TeamsManagementTool.ps1 -RemoveChannel -GroupName "Your Group Name" -ChannelName "Channel Name"
   ```

6. **Restore a Deleted Group:**
   ```powershell
   .\TeamsManagementTool.ps1 -RestoreGroup -GroupName "Your Group Name"
   ```

## Command-line Parameters

- `-CreateGroup`: Create a new Teams group.
- `-CreateChannel`: Create a new channel within a Teams group.
- `-ModifyGroup`: Modify an existing Teams group (e.g., change the group name).
- `-ModifyChannel`: Modify an existing channel within a group.
- `-AddMember`: Add a new member to a Teams group.
- `-RemoveMember`: Remove a member from a Teams group.
- `-DeleteGroup`: Delete an existing Teams group.
- `-DeleteChannel`: Delete a channel from a group.
- `-RestoreGroup`: Restore a previously deleted Teams group.
- `-SetGroupImage`: Set an image for a group.

### Optional Flags

- `-Verbose`: Provides detailed information during script execution.
- `-Force`: Force the execution of certain actions without prompting.

## Error Handling

The script includes error handling for the following situations:
- Invalid group or channel names
- Insufficient permissions
- Group/channel/member not found

If you encounter issues, ensure you have the proper admin rights and that the group or channel names are correct.

## License
This project is licensed and worked under O365 and Exchnage environment.


## Contributions

Contributions are welcome! Please feel free to open an issue or submit a pull request for bug fixes or feature enhancements.

## Contact

For any questions or feedback, please contact the author at **kirtiranjan1988@gmail.com**.

---

Feel free to adjust the placeholders like the repository URL, author contact info, and license as needed. Let me know if you'd like any additional changes!
