'---------------------------------------------------------------------------------------------
' Name          :   ConnectionAPI
' Purpose       :   Creates the Meridian ConnectionAPI object
' Parameters    :   Nothing
' Returns       :   Meridian ConnectionAPI object
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Class ConnectionAPI
	
	Private underlyingObject        
	
    Private Sub Class_Initialize() 
    	AddToDebugLog("Entering Initialize ConnectionAPI, Api = " & IsObject(underlyingObject))   	
		Set underlyingObject = AMCreateObject("ICMeridianAPI.Connection",True)
        AddToDebugLog("Exiting Initialize ConnectionAPI, Api = " & IsObject(underlyingObject))
	End Sub        	
	
    Private Sub Class_Terminate()
    	AddToDebugLog("Entering Terminate ConnectionAPI, Api = " & IsObject(underlyingObject))      	
		underlyingObject.dispose    	
		Set underlyingObject = Nothing
        AddToDebugLog("Exiting  Terminate ConnectionAPI, Api = " & IsObject(underlyingObject))  
	End Sub         	
    
    Public Function SafeObject()
    	AddToDebugLog("Entering SafeObject ConnectionAPI, Api = " & IsObject(underlyingObject))  
    	Set SafeObject = underlyingObject   	
        AddToDebugLog("Exiting SafeObject ConnectionAPI, Api = " & IsObject(underlyingObject))  
    End Function
	
End Class