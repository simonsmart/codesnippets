'---------------------------------------------------------------------------------------------
' Name          :   AddToDebugLog
' Purpose       :   Add text to the debugLog
' Parameters    :   strMessage = String - Message to log
' Returns       :   Nothing
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Sub AddToDebugLog(strMessage)
    If cLog Then
        If gLogFilePath = "" Then 'if the log path is not defined in the script then the following
            gLogFilePath = cLogFolder & Year(Today) & "_" & Month(Today) & "_" & Day(Today) & "_" & Vault.DataStoreName & "_" & User.Name & "_Debug.Log" 
        Else
            'Use the global Log File Path
        End If
        If Not InitializeDebugLogFile(gLogFilePath) Then
           'Debug log text file could not be initialized
        Else
         	On Error Resume Next
            	gLogObjTextFile.WriteLine(Today & "  :  " & strMessage)
            On Error goto 0 
        End If
    Else
        'Logging is disabled
    End If
End Sub

'---------------------------------------------------------------------------------------------
' Name          :   InitializeDebugLogFile
' Purpose       :   Opens or creates the debug log file
' Parameters    :   aFileName = String - File name of the debug log file
' Returns       :   boolean if the file is there or not
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Function InitializeDebugLogFile(aFileName)
    If gLogObjFSO Is Nothing Then
		On Error Resume Next
    	Set gLogObjFSO = AMCreateObject("Scripting.FileSystemObject")
		If gLogObjFSO.FileExists(aFileName) Then 'check if the file exists
			Set gLogObjTextFile = gLogObjFSO.OpenTextFile(aFileName, 8) ' if it exists, append 8=Append
		Else
			Set gLogObjTextFile = gLogObjFSO.CreateTextFile(aFileName, False) 'not exisiting so create file
		End If
    Else
        'FSO has already been used to get the text file object
	End If
	
	If Err <> 0 Then
		InitializeDebugLogFile = False
	Else
		InitializeDebugLogFile = True
	End If
	On Error goto 0
End Function