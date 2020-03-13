'---------------------------------------------------------------------------------------------
' Name          :   CombineArray
' Purpose       :   Combines Two Arrays
' Parameters    :   arr1 = Array - Array 1 to be combined
'               :   arr2 = Array - Array 2 to be combined
'               :   blnRemoveDuplicates = Boolean - Remove duplicates
'               :   blnSort = Boolean - Sort the Combined Array
' Returns       :   A combined array
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Function CombineArray(arr1,arr2,blnRemoveDuplicates,blnSort)
    AddToDebugLog("Entering CombineArray, arr1 = " & Join(arr1) & ", arr2 = " & Join(arr2) &_
        ", blnRemoveDuplicates = " & blnRemoveDuplicates & ", blnSort = " & blnSort)
    Dim objArrayList, strItem
    Set objArrayList = AMCreateObject("System.Collections.ArrayList",True) 'Create arraylist object
     
    'check if it is an array, otherwise make it an array
    If Not IsArray(arr1) Then
        arr1 = Array(arr1)  'always make sure it is an array (string will be transformed to an array)
    End If

    If Not IsArray(arr2) Then 
        arr2 = Array(arr2)'always make sure it is an array (string will be transformed to an array)
    End If
     
    'remove blanks and duplicates
	For Each strItem in arr1 'Loop through each item in the array
		If Not IsNullOrEmpty(strItem) Then 'Remove blanks and nulls
            If blnRemoveDuplicates Then
            	If Not objArrayList.contains(strItem) Then  'Compare if item is already in the Array List
                    objArrayList.Add strItem 'Add Item
                End If
            Else
                objArrayList.Add strItem 'Add Item
            End If
		End If
    Next
     
	For Each strItem in arr2 'Loop through each item in the array
		If Not IsNullOrEmpty(strItem) Then 'Remove blanks and nulls
            If blnRemoveDuplicates Then
            	If Not objArrayList.contains(strItem) Then  'Compare if item is already in the Array List
                    objArrayList.Add strItem 'Add Item
                End If
            Else
                objArrayList.Add strItem 'Add Item
            End If
		End If
    Next
    
    If blnSort Then
        objArrayList.Sort() 'Sort the ArrayList Object
    End If

    CombineArray = objArrayList.ToArray 'Pass array back to function

    Set objArrayList = Nothing 'Clear Object
    AddToDebugLog("Exiting CombineArray, Result = " & Join(CombineArray))
End Function

'---------------------------------------------------------------------------------------------
' Name          :   GetOneDMArray
' Purpose       :   Gets a one dimensional array from a multi dimensional array
' Parameters    :   arrArray = Array - Multi Dimensional Array to get a oneDM array from
'               :   intColumn = Integer - Column number of the MD Array to get (0 is the first)
' Returns       :   A one DM array 
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Function GetOneDMArray(arrArray,intColumn)
    AddToDebugLog("Entering GetOneDMArray, arrArray = " & arrArray & ", intColumn = " & intColumn)
	Dim objArrayList, intCounter

    If Not IsArray(arrArray) Then 
        Exit Function 'if it is not an array exit the function
    End If

    On Error Resume Next
    If ubound(arrArray,2) <> "" Then
        'Continue
    Else
        Exit Function 'not a multi dimensional array, exit
    End If
    If Err.Number <> 0 Then 'If this is not a multidimensional array, error is raised and exit function
        On Error GoTo 0 
        Exit Function
    End If
	
    ReDim arrOneDMArray(ubound(arrArray,2)) 'Create a variable to the maximum amount of rows the column has
   	For intCounter = 0 To ubound(arrArray,2) 'loop through the rows of that column
       	arrOneDMArray(intCounter) = arrArray(intColumn,intCounter) 'add value to array
    Next  

    GetOneDMArray = arrOneDMArray 'Write array back to function
    AddToDebugLog("Exiting GetOneDMArray, Result = " & Join(GetOneDMArray))
End Function

'---------------------------------------------------------------------------------------------
' Name          :   RemoveArray
' Purpose       :   Removes a value from an array
' Parameters    :   arr1 = Array - Array to get items removed from
'               :   arr2 = Array - list to be removed from arr1
' Returns       :   arr1 without arr2's items
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
 Function RemoveArray(arr1,arr2)
    AddToDebugLog("Entering RemoveArray, arr1 = "  & Join(arr1) & ", arr2 = " & Join(arr2))
    Dim strItem
 
    'check if it is an array, otherwise make it an array
    If Not IsArray(arr1) Then 
        arr1 = Array(arr1) 'always make sure it is an array (string will be transformed to an array)
    End if

    If Not IsArray(arr2) Then 
        arr2 = Array(arr2) 'always make sure it is an array (string will be transformed to an array)
    End if
     
    'filters every value in the second array from the first array.
    For each strItem In arr2 'loop array of to remove items
        If Not IsNullOrEmpty(strItem) Then 'check if not is null or empty value
            arr1 = Filter(arr1,strItem,0,1) 'Pick strItem out of arr1
        End If
    Next
         
    RemoveArray = arr1
    AddToDebugLog("Exiting RemoveArray, Result = " & Join(RemoveArray))
End Function

'---------------------------------------------------------------------------------------------
' Name          :   SortArray
' Purpose       :   Sorts an Array
' Parameters    :   arrArray = Array - Array that needs to be sorted
'               :   blnAddBlankAtTop = Boolean - Add a blank at the top of the array
'               :   blnRemoveDuplicates = Boolean - Remove duplicates
' Returns       :   A sorted array 
' 
' Change Log
' Date              Programmer      Description
' 18-10-2016        Simon Smart      First Version
'---------------------------------------------------------------------------------------------
Function SortArray(arrArray,blnAddBlankAtTop,blnRemoveDuplicates)
    AddToDebugLog("Entering SortArray, arrArray = " & Join(arrArray) & ", blnAddBlankAtTop = " & blnAddBlankAtTop & ", blnRemoveDuplicates = " & blnRemoveDuplicates)
	Dim objArrayList, strItem
	Set objArrayList = AMCreateObject("System.Collections.ArrayList",True) 'Create Array List Object
	
	For Each strItem in arrArray 'Loop through each item in the array
		If Not IsNullOrEmpty(strItem) Then 'Check if not is null or empty string
            If blnRemoveDuplicates Then
            	If Not objArrayList.contains(strItem) Then  'Compare if item is already in the Array List
                    objArrayList.Add strItem 'Add Item
                End If
            Else
                objArrayList.Add strItem 'Add Item
            End If
		End If
    Next
	
    If blnAddBlankAtTop = True Then 
        objArrayList.Add "" 'Add blank
    End if
    
    objArrayList.Sort() 'Sort the ArrayList Object
	SortArray = objArrayList.ToArray 'Pass array back to function

	Set objArrayList = Nothing 'Clear Object
	AddToDebugLog("Exiting SortArray, Result = " & Join(SortArray))
End Function