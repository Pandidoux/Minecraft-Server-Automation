# Minecraft-Server-Automation
Minecraft-Server-Automation is a Visual Basic Script created to automate tasks for your Minecraft server.

## Features
- Stop your server if no player was online since a certain amount of time.
- Backup some configurable files and folders after while server is running (Save-Backup). You can choose to do it only if there are some player online to save disk space and avoid useless.
- Backup your entire server folder after an automatic stop by script (Dump-Backup)
- Reboot the server if it's not running anymore. (Verification method use server query or window detection)
- Compress your backups with 7zip
- Log everything  in a logfile.
- Send commands to the server via server windows. (I planned to use server rcon in a future update)
- Easy development (vbs) if you want to program your custom actions.

## Installation
### 1) Basic install
Just clone the repository :<br>
`git clone https://github.com/Pandidoux/Minecraft-Server-Automation.git`
<br>**OR**<br>
Download the zip : [Minecraft-Server-Automation-main.zip](https://github.com/Pandidoux/Minecraft-Server-Automation/archive/refs/heads/main.zip).
<br><br>**AND**<br>
Place its contents **outside** of your Minecraft server folder.

### 2) Full install (recommended)
This script use the Minecraft server query to enable some features.<br>
You can find the Minecraft-Server-Query repository [here](https://github.com/Pandidoux/Minecraft-Server-Query.git).<br>
Download the lastest release of the Minecraft-Server-Query.jar [here](https://github.com/Pandidoux/Minecraft-Server-Query/releases).<br>
Then place the Minecraft-Server-Query.jar in the root directory of the script.

## Usage
- To start the script, execute `Server-Automation_Start.bat`.
- To configure the script, edit `Server-Automation.vbs` with a text editor.
### Configuration
To configure the script for your server you have to edit some User Variables in the begining of the `Server-Automation.vbs` file.<br>
Some comment are here to help you understand what to do with each variables.<br>
Each comment begin with the apostrophe `'` character.<br>
/!\ Make sure to fill every uncompleted path correctly /!\
/!\ Fill every parameters, or use default value /!\

#### Sections :
- **Script Parameters**: Parameters used for all general purpose in the script
- **Query Parameters**: Parameters used for the server query configuration (ennable your `enable-query` in your `server.properties`)
- **Custom Actions**: Parameters used to ennable custom actions if you want to create your own.
- **Stop Server**: Parameters used to configure the automatic /stop of your server.
- **Server Dump**: Parameters used to configure the entire backup of the server folder.
- **Server Save**: Parameters used for the periodic save backup of the server and after automatic /stop
- **Reboot server**: Parameters used for rebooting the server a crash is detected.
- **7zip Parameter**: Parameters used for backup compression.
