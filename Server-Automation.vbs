Option Explicit
Dim stopServer_ExecutionInterval,stopServer_StopDelay,server_WindowName,rebootServer_RetryNumber,saveBackup_SaveInterval,script_RefreshInterval,server_SaveAllDelay,server_DelayAfterRestart,query_IpAdress,query_Port,query_JarFile,query_OutputFile,query_OutputFormat,customAction_ExecutionInterval,stopServer_StopIfNobodyOnline,stopServer_MaxTimeNobodyOnline,stopServer_ShutdownComputer,dumpBackup_DoBackupAfterStop,dumpBackup_ArchivesDestination,saveBackup_DoSaveAfterStop,saveServerEvery,saveBackup_SaveIntervalOnlyIfPlayerOnline,saveBackup_SaveList,saveBackup_ArchiveDestination,rebootServer_RebootIfNotRunning,rebootServer_VerficationInterval,server_StartScript,rebootServer_BeforeRestartDelay,server_ServerFolder,rebootServer_VerificationMethod,exe7zip,compression7zip,minecraftQueryOutputFile
'____________________User Variables Begin (Edit this)____________________
' Script Parameters (Values must be correctly set no matter what)
script_RefreshInterval = 1 ' Check every X seconds if we have to do someting
server_WindowName = "/usr/bin/bash --login -i C:\Users\...\Server\Start.sh" ' Commands will be send to this window
server_SaveAllDelay = 30 ' Wait X seconds after sending /save-all to wait for the server to finish
server_ServerFolder = "C:/Users/.../Server" ' Path to the server folder

' Query Parameters
query_IpAdress = "127.0.0.1" ' IP address of the server
query_Port = "25565" ' Query port of the server
query_JarFile = "Minecraft-Server-Query.jar" ' Edit if renamed
query_OutputFile = "Minecraft-Server-Query.txt" ' Edit if renamed
query_OutputFormat = "RAW_BASIC" ' Output format of query_JarFile

' Custom Actions
customAction_ExecutionInterval = 0 ' Do your custom action on server every X seconds, O to disable

' Stop Server (Need Query Parameters)
stopServer_StopIfNobodyOnline = TRUE ' Ennable server shutdown if no player online for more than [stopServer_MaxTimeNobodyOnline] seconds
stopServer_ExecutionInterval = 10 ' Verification interval (in seconds) for server stop
stopServer_MaxTimeNobodyOnline = 30 ' Stop the server if no player are online since X minutes
stopServer_ShutdownComputer = FALSE ' Enable shutdown after [stopServer_MaxTimeNobodyOnline] seconds if nobody is online
stopServer_StopDelay = 30 ' Wait X seconds after sending /stop to wait for the server to finish

' Server Dump, Backup the entire folder of the server
dumpBackup_DoBackupAfterStop = FALSE ' Backup server folder after automatic /stop
dumpBackup_ArchivesDestination = "C:/Users/.../ServerDump" ' Backup files will be saved in this directory

' Server Save, Backup specific folder of the server
saveBackup_DoSaveAfterStop = FALSE ' Backup specific files after automatic /stop
saveBackup_SaveInterval = 0 ' Save specific files every X minutes while server is running, 0 to disable
saveBackup_SaveIntervalOnlyIfPlayerOnline = TRUE ' Save specific files every [saveBackup_SaveInterval] seconds only if somebody is online (need Query Parameters)
saveBackup_SaveList = Array("logs","plugins","world","world_nether","world_the_end","banned-ips.json","banned-players.json","ops.json","server.properties","Start.sh","whitelist.json") ' Name of the folders and files to include inside the save
saveBackup_ArchiveDestination = "C:/Users/.../ServerSave" ' Backup files will be saved in this directory

' Reboot server
rebootServer_RebootIfNotRunning = FALSE ' Reboot server if cannot be found
rebootServer_VerficationInterval = 30 ' Verification interval in seconds if server is up. (The server window will be put in forground, keep value high)
server_StartScript = "Start.sh" ' Name of the start script of the server (ex: start.bat, start.sh)
rebootServer_BeforeRestartDelay = 10 ' Wait X seconds BEFORE restarting the server
server_DelayAfterRestart = 60 ' Wait for server to finish restarting, delay in seconds
rebootServer_RetryNumber = 5 ' Retry X time to reach the server, then restart the server
rebootServer_VerificationMethod = "USE_QUERY" ' Method to verify if server is running :
' => "USE_WINDOW" (Need [server_WindowName], only localhost) '
' => "USE_QUERY" (Need all Query Parameters)

' 7zip Parameter (Edit if you use Dump or Save)
exe7zip = "C:/Program Files/7-Zip/7z.exe" ' Path to the 7z.exe file
compression7zip = 9 ' Compression 0=none, 1=fast, 3=low, 5=normal, 7=high, 9=maximum

'____________________User Variables End____________________


'____________________Do not edit____________________
Dim debuging,StartTime,EndTime,fso,WShell
StartTime = Timer()
debuging = false
Set fso = CreateObject("Scripting.FileSystemObject")
Set WShell = CreateObject("WScript.Shell")
minecraftQueryOutputFile = getCurrentDirectory() & "/" & query_OutputFile
'____________________Functions Begin____________________
Function log(textLineToLog) ' NOK
    Wscript.Echo textLineToLog
    Dim logFilePath, logFileObj
    logFilePath = Replace(WScript.CreateObject("WScript.Shell").currentDirectory,"\","/") & "\Server-Automation.log"
    If NOT(fso.FileExists(logFilePath)) Then
        Call fso.CreateTextFile(logFilePath,False) 'Create logFile, no replace, Unicode
    End If
    Set logFileObj = fso.OpenTextFile(logFilePath,8)
    logFileObj.WriteLine("[" & getDateTime() & "] " & textLineToLog)
    logFileObj.Close()
    Set logFilePath = Nothing
End Function ' log
Function getCurrentDirectory() ' Return current directory of the calling script
    getCurrentDirectory = Replace(WScript.CreateObject("WScript.Shell").currentDirectory,"\","/")
end Function
Function println(text) ' Print text and go next line
    Wscript.Echo text
end Function
Function sleep(ms) ' pause program for X milliseconds
    WScript.Sleep ms
End Function
Function wait(seconds) ' Pause script for X seconds
    WScript.Sleep seconds*1000
end Function
Function getDateTime() ' Return date format "YYYY-MM-DD HH:mm:ss"
    getDateTime = Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2) & " " & Time()
end Function
Function getDate() ' Return date format "YYYY-MM-DD"
    getDate = Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2)
end Function
Function getTime() ' Return date format "HH-mm-ss"
    getTime = Replace(Time(),":","-")
end Function
Function read_line(filepath,n) ' Read line n of the file filepath and return text
    Dim objFile, arrLines
	Set objFile = fso.OpenTextFile(filepath,1)
	arrLines = Split(objFile.ReadAll,vbCrLf)
	If UBound(arrLines) >= n-1 Then
		If arrLines(n-1) <> "" Then
			read_line = arrLines(n-1)
		Else
			read_line = "Line " & n & " is null."
		End If
	Else
		read_line = "Line " & n & " does not exist."
	End If
	objFile.Close
	Set objFile = Nothing
    Set arrLines = Nothing
End Function
Function getPlayerNumber() ' Return online player number using the query output file
    Dim queryOutput, lineItemsArray
    queryOutput = read_line(minecraftQueryOutputFile,1)
    lineItemsArray = Split(queryOutput,", ")
    getPlayerNumber = lineItemsArray(3)
    Set queryOutput = Nothing
    Set lineItemsArray = Nothing
end Function
Function sendCommandToWindow(command,windowName)
    WShell.AppActivate(windowName)
    wshell.SendKeys(command)
    wshell.SendKeys("{ENTER}")
end Function
Function fileDelete(fileToDelete) ' Delete a file, return 1 if success else 0
    If fso.FileExists(fileToDelete) Then
        fso.deleteFile(fileToDelete)
        fileDelete = 1
    Else
        fileDelete = 0
    End If
end Function
Function dumpServer() ' Dump [server_ServerFolder] to [dumpBackup_ArchivesDestination]
    Dim zipName, zipCommand
    zipName = "Dump_" & getDate() & "_" & getTime() & ".7z"
    zipCommand = chr(34) & exe7zip & chr(34) & " a -t7z " & chr(34) & dumpBackup_ArchivesDestination & "/" & zipName & chr(34) & chr(32) & chr(34) & server_ServerFolder & chr(34) & " -mx" & compression7zip & " -mmt=on"
    log("[INFO] Dumping server folder... " & chr(34) & server_ServerFolder & chr(34))
    Call WShell.Run(zipCommand,1,true)
    log("[INFO] Dump archive saved to: " & chr(34) & dumpBackup_ArchivesDestination & "/" & zipName & chr(34))
    Set zipName = Nothing
    Set zipCommand = Nothing
End Function
Function saveServer() ' Save specific files and folder [saveFile] from [server_ServerFolder] to [saveBackup_ArchiveDestination]
    Dim allItemsPath, allItems, item, zipName, zipCommand
    For Each item In saveBackup_SaveList
        If allItemsPath <> "" Then
            allItemsPath = allItemsPath & chr(32)
            allItems = allItems & ", "
        End If
        allItemsPath = allItemsPath & chr(34) & server_ServerFolder & "/" & item & chr(34)
        allItems = allItems & chr(34) & item & chr(34)
    Next
    zipName = "Save_" & getDate() & "_" & getTime() & ".7z"
    zipCommand = chr(34) & exe7zip & chr(34) & " a -t7z " & chr(34) & saveBackup_ArchiveDestination & "/" & zipName & chr(34) & " " & allItemsPath & " -mx" & compression7zip & " -mmt=on"
    log("[INFO] Saving server files...")
    Call WShell.Run(zipCommand,1,true)
    log("[INFO] Save archive saved to: " & chr(34) & saveBackup_ArchiveDestination & "/" & zipName & chr(34))
    Set allItemsPath = Nothing
    Set allItems = Nothing
    Set item = Nothing
    Set zipName = Nothing
    Set zipCommand = Nothing
End Function
Function windowExists(txtWindowTitle) ' Use AppActivate to see if the window exists (return 1 if exists else 0)
    Dim ret
    ret = WShell.AppActivate(txtWindowTitle) 
    If ret = True Then 
        windowExists = 1
    Else
        windowExists = 0
    End If
    Set ret = Nothing
End Function
Function isFileExists(filePathToCheck) ' If File is reachable return 1, Else return 0
    If fso.FileExists(filePathToCheck) Then
        isFileExists = 1
    Else
        isFileExists = 0
    End If
End Function
Function doServerQuery() ' Start jar file to query, return 1 if query OK else false
    Dim command
    command = "java -jar " & chr(34) & getCurrentDirectory() & "/" & query_JarFile & chr(34) & " -IP " & query_IpAdress & " -PORT " & query_Port & " -FORMAT " & query_OutputFormat & " -OUTPUT " & chr(34) & query_OutputFile & chr(34)
    WShell.Run command
    doServerQuery = isFileExists(minecraftQueryOutputFile)
    Set command = Nothing
end Function
Function serverStart()
    Dim startCommand
    If rebootServer_BeforeRestartDelay > 0 Then
        log("[INFO] Wait " & rebootServer_BeforeRestartDelay & " seconds before starting the server...")
        wait(rebootServer_BeforeRestartDelay)
    End If
    log("[INFO] Restarting server...")
    startCommand = chr(34) & server_ServerFolder & "/" & server_StartScript & chr(34)
    WShell.CurrentDirectory = server_ServerFolder
    Call WShell.Run(startCommand,1,false)
    If server_DelayAfterRestart > 0 Then
        log("[INFO] Wait " & server_DelayAfterRestart & " seconds for the server to start...")
        wait(server_DelayAfterRestart)
    End If
    Set startCommand = Nothing
End Function
'____________________Functions End____________________


'____________________ Main Actions ____________________
Dim playerNumber,retryCount,isServerUp,secondsEllapsedSinceNobodyOnline,timeLeftBeforeShutdown,resQuery,resDelete,timeElapsedSinceLastQuery,timeElapsedSinceLastStopServer,timeElapsedSinceLastCustomAction,timeElapsedSinceLastSaveServerEvery,timeElapsedSinceLastServerUpVerification
secondsEllapsedSinceNobodyOnline = 0
timeLeftBeforeShutdown = stopServer_MaxTimeNobodyOnline*60
timeElapsedSinceLastQuery = 0
If rebootServer_RetryNumber < 1 Then
    rebootServer_RetryNumber = 1
End If

' Log start info
log("[INFO] ----------START----------")
If debuging Then
    log("[WARN] Debunging mode is ennable, turn off debuging to use the script")
End If
If customAction_ExecutionInterval > 0 Then
    log("[INFO] Do custom action every " & customAction_ExecutionInterval & " seconds")
End If
If stopServer_StopIfNobodyOnline Then
    log("[INFO] Stop server after " & stopServer_MaxTimeNobodyOnline & " minutes without player online, verification every " & stopServer_ExecutionInterval & " seconds")
End If
If stopServer_ShutdownComputer Then
    log("[INFO] Shutdown computer after " & stopServer_MaxTimeNobodyOnline & " minutes without player online")
End If
If saveBackup_DoSaveAfterStop Then
    log("[INFO] Make a Save backup when server stop")
End If
If dumpBackup_DoBackupAfterStop Then
    log("[INFO] Make a Dump backup of the server folder when server stop")
End If
If saveBackup_SaveInterval > 0 Then
    Dim onlyOnline
    If saveBackup_SaveIntervalOnlyIfPlayerOnline Then
        onlyOnline = " only if there are player online"
    Else
        onlyOnline = ""
    End If
    log("[INFO] Make a Save backup every " & saveBackup_SaveInterval & " minutes" & onlyOnline)
End If
If rebootServer_RebootIfNotRunning Then
    log("[INFO] Reboot server if window cannot be found, verify every " & rebootServer_VerficationInterval & " seconds")
End If


' Test here your custom code, command etc...
If debuging Then
    ' This part of code will be executed only one time
    
End If



Do While NOT debuging ' Loop checking every script_RefreshInterval seconds
    wait(script_RefreshInterval)


    ' ---------- Custom Action Begin ----------
    If customAction_ExecutionInterval <> 0 Then
        If timeElapsedSinceLastCustomAction >= customAction_ExecutionInterval Then
            ' Do some actions here...
            log("[INFO] Execute custom actions")
            Call sendCommandToWindow("say Hello everyone !",server_WindowName) ' Say hello
            timeElapsedSinceLastCustomAction = 0 ' Reset timer
        Else
            timeElapsedSinceLastCustomAction = timeElapsedSinceLastCustomAction + script_RefreshInterval ' Increment timer
        End If
    End If
    ' ---------- Custom Action End ----------


    ' ---------- Save Backup Server Every X second Begin ----------
    If saveBackup_SaveInterval*60 <> 0 Then
        If timeElapsedSinceLastSaveServerEvery >= saveBackup_SaveInterval*60 Then
            If saveBackup_SaveIntervalOnlyIfPlayerOnline Then
                log("[INFO] Save-Server: Do server query...")
                resQuery = doServerQuery()
                If resQuery = 1 Then
                    playerNumber = getPlayerNumber()
                    resDelete = fileDelete(minecraftQueryOutputFile)
                    If resDelete = 0 Then
                        log("[ERROR] Save-Server: Cannot delete file: " & minecraftQueryOutputFile)
                    End If
                    If playerNumber > 0 Then
                        log("[INFO] " & playerNumber & " players online, sending command /save-all to the server")
                        Call sendCommandToWindow("save-all",server_WindowName) ' Save server
                        println("[INFO] Wait " & server_SaveAllDelay & " seconds for server to finish save-all...")
                        wait(server_SaveAllDelay) ' Wait for server to save-all
                        Call saveServer()
                    Else
                        log("[INFO] Save-Server: No players online on the server, save cancelled")
                    End If
                Else
                    log("[ERROR] Save-Server: Query failed")
                    timeElapsedSinceLastSaveServerEvery = 0 ' Reset timer
                End If
            Else
                log("[INFO] Sending command /save-all to the server")
                Call sendCommandToWindow("save-all",server_WindowName) ' Save server
                println("[INFO] Wait " & server_SaveAllDelay & " seconds for server to finish save-all...")
                wait(server_SaveAllDelay) ' Wait for server to save-all
                Call saveServer()
            End If
            timeElapsedSinceLastSaveServerEvery = 0 ' Reset timer
        Else
            timeElapsedSinceLastSaveServerEvery = timeElapsedSinceLastSaveServerEvery + script_RefreshInterval ' Increment timer
        End If
    End If
    ' ---------- Save Backup Server Every X second End ----------


    ' ---------- Reboot Server Begin ---------- 
    If rebootServer_RebootIfNotRunning Then
        If timeElapsedSinceLastServerUpVerification >= rebootServer_VerficationInterval Then
            If rebootServer_VerificationMethod = "USE_WINDOW" Then
                For retryCount=1 To rebootServer_RetryNumber
                    isServerUp = windowExists(server_WindowName)
                    If isServerUp = 0 Then
                        log("[WARN] Reboot-Server: Attenpt(" & retryCount & ") Server unreachable ! Wait 10 seconds and retry")
                        wait(10)
                    Else
                        ' If one attept is successful, assume server is up
                        log("[INFO] Reboot-Server: Attenpt(" & retryCount & ") Server is up")
                        Exit For
                    End If
                Next
                If isServerUp = 0 Then
                    ' Exit for loop after X retry with server down
                    log("[ERROR] Reboot-Server: Server crash ! Server window not found.")
                    log("[INFO] Reboot-Server: Reboot the server...")
                    serverStart()
                End If
            End If
            If rebootServer_VerificationMethod = "USE_QUERY" Then
                For retryCount=1 To rebootServer_RetryNumber
                    log("[INFO] Reboot-Server: Do server query...")
                    isServerUp = doServerQuery()
                    If isServerUp = 0 Then
                        log("[WARN] Reboot-Server: Attenpt(" & retryCount & ") Server unreachable ! Wait 10 seconds and retry")
                        wait(10)
                    Else
                        ' If one attept is successful, assume server is up
                        log("[INFO] Reboot-Server: Attenpt(" & retryCount & ") Server is up")
                        Exit For
                    End If
                Next
                If isServerUp = 0 Then
                    ' Exit for loop after X retry with server down
                    log("[ERROR] Reboot-Server: Server crash ! Query unable to reach the server.")
                    log("[INFO] Reboot-Server: Rebooting the server...")
                Else
                    resDelete = fileDelete(minecraftQueryOutputFile)
                    If resDelete = 0 Then
                        log("[ERROR] Reboot-Server: Cannot delete file: " & minecraftQueryOutputFile)
                    End If
                End If
            End If
            timeElapsedSinceLastServerUpVerification = 0 ' Reset timer
        Else
            timeElapsedSinceLastServerUpVerification = timeElapsedSinceLastServerUpVerification + script_RefreshInterval ' Increment timer
        End If
    End If
    ' ---------- Reboot Server End ----------


    ' ---------- Stop and Shutdown Begin ----------
    If stopServer_StopIfNobodyOnline AND stopServer_ExecutionInterval > 0 Then
        If timeElapsedSinceLastStopServer >= stopServer_ExecutionInterval Then
            log("[INFO] Stop-Server: Do server query...")
            resQuery = doServerQuery()
            If resQuery = 1 Then
                playerNumber = getPlayerNumber()
                resDelete = fileDelete(minecraftQueryOutputFile)
                If resDelete = 0 Then
                    log("[ERROR] Stop-Server: Cannot delete file: " & minecraftQueryOutputFile)
                End If
                If playerNumber > 0 Then
                    ' At least one player online
                    log("[INFO] Stop-Server: " & playerNumber & " players online on the server, reset countdown")
                    timeLeftBeforeShutdown = stopServer_MaxTimeNobodyOnline*60 ' Reset timer
                Else
                    ' No player online
                    timeLeftBeforeShutdown = timeLeftBeforeShutdown - stopServer_ExecutionInterval ' Decrement timer
                    log("[INFO] Stop-Server: No players online on the server, " & timeLeftBeforeShutdown & " seconds left before server stop")
                    If timeLeftBeforeShutdown <= 0 Then
                        ' Timer as reach the limit, stop server
                        ' Sending command SAVE-ALL to server
                        log("[INFO] Stop-Server: Sending command /save-all to the server")
                        Call sendCommandToWindow("save-all",server_WindowName) ' Save server
                        println("[INFO] Stop-Server: Wait " & server_SaveAllDelay & " seconds for server to finish save-all...")
                        wait(server_SaveAllDelay) ' Wait for server to save-all
                        ' Sending command STOP to server
                        log("[INFO] Stop-Server: Sending command /stop to the server")
                        Call sendCommandToWindow("stop",server_WindowName) ' Stop server
                        println("[INFO] Stop-Server: Wait 30 seconds for server to be stoped...")
                        wait(30) ' Wait for server to be stoped
                        If dumpBackup_DoBackupAfterStop Then
                            ' Dump server folder
                            Call dumpServer()
                        End if
                        If saveBackup_DoSaveAfterStop Then
                            ' Save some files and folders of the server
                            Call saveServer()
                        End if
                        If stopServer_ShutdownComputer Then ' Shutdown computer
                            log("[INFO] Stop-Server: Shutdown the computer in 2 minutes, Bye bye !!!")
                            WShell.Run("shutdown -s -t 120")
                        End If
                        Exit Do ' Exit loop
                    ' Else
                    '     log("[INFO] Stop-Server: Server will stop in " & timeLeftBeforeShutdown & " seconds")
                    End If
                End If
            Else
                log("[ERROR] Server-Stop: Query output file cannot be found")
            End If
            timeElapsedSinceLastStopServer = 0 ' Reset timer
        Else
            timeElapsedSinceLastStopServer = timeElapsedSinceLastStopServer + script_RefreshInterval ' Increment timer
        End If
    End If



Loop ' End Loop checking every script_RefreshInterval seconds

' End script
EndTime = Timer()
log("[INFO] End of script, exiting now after " & FormatNumber(EndTime - StartTime, 2) & " seconds of execution")
log("[INFO] ----------END----------")
