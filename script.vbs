'____________________________________________________________________________________________________________
'
' PROGRAM TO SYNC PRONOTE WITH ACTIVE DIRECTORY
'
' MADE BY VICTOR BOISSSIERE WITH THE HELP OF MICHEL PORTAL
'
' DATE : 10 Jun 2015
'
'____________________________________________________________________________________________________________


' PARAMETERS - You can change this data when you need it
'-------------------------------------------------------------------------------------------------------------

Set fso = CreateObject("Scripting.FileSystemObject")
xmlpath = fso.BuildPath("C:\Users\vboissiere\Google Drive", "\" & "index.xml")
myLdapPath = "dc=claudel,dc=lan"

userFriendlyOldDirectory = "Utilisateurs/Anciens Eleves"

'-------------------------------------------------------------------------------------------------------------
' END PARAMETERS


' MAIN PROGRAM
'-------------------------------------------------------------------------------------------------------------

WScript.Echo "SYNCING PROGRAM FOR PRONOTE -- Choose options"

'Main loop of the program
Do While True

	'Display the menu
	WScript.Echo vbLf & "1. Sync" & vbLf & "2. Display index" & vbLf & "3. Add index" & vbLf & "4. Remove index" & vbLf & "0. Exit"
	
	choice = askInputNumber()
	
	'Check option and trigger corresponding functions
	If choice >= 0 Then
		
		'Each function verify that the configuration exists in order to run
		Select Case True
			'Check configuration exists else ask user if he wants to create it
			Case Not configurationExists()
			Case choice = 1
				sync()
			Case choice = 2
				'Display table
				displayConfiguration()
			Case choice = 3
				'Ask user data and add it if valid to the XML configuration file
				writeConfiguration()
			Case choice = 4
				'Ask user ID of index and remove it from the XML configuration file
				removeConfiguration()
			Case choice = 0
				WScript.Quit
			Case Else
				displayError("Not valid")
		End Select
		
	Else
		displayError("Not a number")
	End If
	
Loop

'-------------------------------------------------------------------------------------------------------------
' END MAIN PROGRAM


'SYNCING FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

'Sync active directory with pronote data
Sub sync()
	'Get all active Directory paths in the XML configuration file
	Set indexes = getXMLIndexes()
	
	'Add path for old Students directory
	
	'Old Students directory pronote index is empty (that is how we can recognize it)
	
	WScript.Echo "Indexing Active Directory...."
	Set activeDirectoryGUID = getActiveDirectoryGUID(indexes.Item("uniqueActiveDirectory"))
	WScript.Echo "Done!"
	
	'Getting needed objects for the search
	dtStart = TimeValue(Now())
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	
	'TODO : Verify that category is indexed in the configuration, use .contains()
	 
	'search each students
	Call searchStudent(objCommand, "eric moen", indexes)
	
	'Close connection
	objConnection.Close
	
End Sub

Sub createStudent()

	WScript.Echo "Creating student... (simulation)"
	
End Sub

'Trigger actions based on the Excel data and the position of the student
Sub studentExists(pronoteExcelIndex, indexes, posFound)

	WScript.Echo "Student Exist..."
	
	WScript.Echo "Trigger corresponding actions... (simulation)"
	
	pronoteADIndex = indexes.Item("pronote").Item(posFound)
	
	If pronoteExcelIndex = pronoteAdIndex Then
		WScript.Echo "Good Section"
	Else
		WScript.Echo "Wrong section"
	End If
	
End Sub

'Search student. Can be used many times
Sub searchStudent(objCommand, studentName, indexes)

	'Search students in the given configuration indexes (AD paths)
	For i = 0 To indexes.Item("uniqueActiveDirectory").Count - 1 Step 1
		
		'Query for the search
		objCommand.CommandText = _
		    "<" & getLdapPath(indexes.Item("uniqueActiveDirectory").Item(i)) & _
		     ">;(&(objectCategory=person)(objectClass=user)(name=" & studentName & "));samAccountName;subtree"
		  
		Set objRecordSet = objCommand.Execute
		 
		'if student found, stop search and execute subRoutine
		If objRecordset.RecordCount <> 0 Then
		    Call studentExists("not set", indexes, i)
		    Exit Sub
		End If
	Next

	'No student found, create student
	Call createStudent()
	
End Sub


'-------------------------------------------------------------------------------------------------------------
' END SYNCING FUNCTIONS




' ACTIVE DIRECTORY FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

'get ldap path based on the path written in the configuration XML file
Function getLdapPath(path)

	'The path should be written as something like "Utilisateurs/Eleves/Eleves du CM2
	OUarray = Split(path,"/")
	
	'Get all OU, with the right slash
	ouPath = ""
	
	'concatenate in reverse order for nicer software (more user friendly)
	For Each x In OUarray
		ouPath = "ou=" & x & "," & ouPath
	Next
	
	getLdapPath = "LDAP://" & ouPath & myLdapPath

End Function

'Get the folder Object to a Active Directory folder. Return Nothing if path is false
Function getActiveOUDirectory(path)

	fullPath = getLdapPath(path)
	
	If fullPath <> "" Then
	
		'Check if OU exists, return the OU object if that is the case, else return null
		On Error Resume Next
		Dim OUobject
		Set OUobject = GetObject(fullPath)
		
		'Display corresponding errors, return null if error(s) else return OU
		Select Case Err.Number
		Case 0
		    Set getActiveOUDirectory = OUobject
		Case &h80072030
		    displayError("OU doesn't exist" & vbLf & vbLf & "Full path : " & fullPath & vbLf)
		    Set getActiveOUDirectory = Nothing
		Case Else
		    displayError("Adding OU failed because OU not valid. Error code : "& Err.Number & vbLf & vbLf & "Full path : " & fullPath & vbLf)
		    Set getActiveOUDirectory = Nothing
		End Select
	Else
		displayError("Wrong format path to active Directory")
		Set getActiveOUDirectory = Nothing
	End If
	

End Function

'Check if data given to add the the xml file is valid
Function validateIndex(pronote, activeDirectory)
	
	If Len(pronote) <= 1 or Len(activeDirectory) <= 1 Then
		displayError("Lenght of Pronote index and active directory index should be both greather than 1")
		validateIndex = False
	End If
	
	validateIndex = True

End Function

'Return an array of the user IDs with the corresponding path in the configuration file
Function getActiveDirectoryGUID(paths)
	
	'Create the list
	Set activeDirectoryGUID = CreateObject("System.Collections.ArrayList")
	
	'Foreach all AD paths in the XML configuration file
	For i = 0 To paths.Count - 1 Step 1
		
		'Get OU
		Set OUs = getActiveOUDirectory(paths.Item(i))
		
		'Check if valid and add unique GUID to the list
		If Not OUs Is Nothing Then
			For Each user in OUs
    			activeDirectoryGUID.Add user.GUID
			Next
		Else
			removeInConfiguration(i)
			displayError("Wrong path : " & paths.Item(i) & "  . This item has been removed.")
		End If
	Next
	
	Set getActiveDirectoryGUID = activeDirectoryGUID
	
End Function

'-------------------------------------------------------------------------------------------------------------
' END ACTIVE DIRECTORY FUNCTIONS




' CONFIGURATION FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

'Display the current configuration of the xml file that contains the settings
Sub displayConfiguration()

	'Load XML
	Set xmlDoc = loadXML()
	
	'Create lists to display in the table
	Set pronote = CreateObject("System.Collections.ArrayList")
	Set activeDirectory = CreateObject("System.Collections.ArrayList")
	
	'Adding MENU ITEM
	pronote.Add "Pronote index"
	activeDirectory.Add "Active directory path"
	
	'Variables to render the correct lenght for the table
	Dim maxPronoteLenght
	maxPronoteLenght = Len(pronote.item(0))
	maxActiveDirectoryLength = Len(activeDirectory.item(0))
	
	'Get nodes
	Set nodes = xmlDoc.documentElement.SelectNodes("//Index")
	
	'If none, display message and exit function
	If nodes.length = 0 Then
		WScript.Echo "There is no index at the time"
		Exit Sub
	End If
	
	'Foreach XML file
	For Each Index In nodes
	
		'Get tag name in the XML
		p = Index.getElementsByTagName("Pronote")(0).text
		a = Index.getElementsByTagName("ActiveDirectory")(0).text
		
		'Display spacing for the table
		maxPronoteLength = max(maxPronoteLenght, Len(p))
		maxActiveDirectoryLength = max(maxActiveDirectoryLength ,Len(a))
		
		'Add elements to the corresponding list
		pronote.Add p
		activeDirectory.Add a
	Next
	
	'Generate top and bottom border
	bottomAndTopBorder = " " & generateTextBorder(maxPronoteLength + maxActiveDirectoryLength + 10)
	
	'Top Border
	WScript.Echo bottomAndTopBorder
	
	For i = 0 To pronote.Count - 1 Step 1
	
		If i = 0 Then 
			id = "id"
		Else
			id = CStr(i)
		End If
		
		'handle spacing
		spaceCol1 = getTableSpace(id, max(pronote.Count / 10 + 2, Len(id)))
		spaceCol2 = getTableSpace(pronote.Item(i), maxPronoteLength)
		spaceCol3 = getTableSpace(activeDirectory.Item(i), maxActiveDirectoryLength)
	
		WScript.Echo "| " & id & spaceCol1 & " | " & pronote.Item(i) & spaceCol2 & " | " & activeDirectory.Item(i) & spaceCol3 & " |"
		
		If i = 0 Then WScript.Echo bottomAndTopBorder End If
		
	Next
	
	'Bottom Border
	WScript.Echo bottomAndTopBorder

	
End Sub

'Write into the current configuration of the xml file that contains the settings
Sub writeConfiguration()

	'Ask the user the right configuration
	WScript.Echo "Name of the index in Pronote : "
	pronote = WScript.StdIn.ReadLine
	
	WScript.Echo vbLf & "Corresponding path in Active Directory : "
	activeDirectory = WScript.StdIn.ReadLine
	
	WScript.Echo vbLf
	
	'Validate data
	If validateIndex(pronote, activeDirectory) Then
		
		Set OUdirectory = getActiveOUDirectory(activeDirectory)
		
		'Check that OU directory is valid (so function does not return null)
		If Not OUdirectory Is Nothing Then
			
			'Adding directory to the XML file
			Set xmlDoc = _
			CreateObject("Microsoft.XMLDOM")
			
			xmlDoc.Async = "False"
			xmlDoc.Load(xmlpath)
			
			Set objRoot = xmlDoc.documentElement
			  
			Set objRecord = _
			  xmlDoc.createElement("Index")
			objRoot.appendChild objRecord
			
			Set objFieldValue = _
			  xmlDoc.createElement("Pronote")
			objFieldValue.Text = pronote
			objRecord.appendChild objFieldValue
			
			Set objFieldValue = _
			  xmlDoc.createElement("ActiveDirectory")
			objFieldValue.Text = activeDirectory
			objRecord.appendChild objFieldValue
			  
			'Save the file
			xmlDoc.Save xmlpath
			
			'Notify user
			WScript.Echo "Added"

		End If
	End If
				
End Sub

'Ask user which index he wants to remove
Sub removeConfiguration()

	WScript.Echo "ID of the index to delete (0 to cancel) : "
	number = askInputNumber()
	
	'The user has cancelled, or wrong input
	If number = 0 Then
		WScript.Echo "Canceled"
	Else If number = -1 Then
			displayError("This is not a number")
		Else
			'Data valid, trigger delete
			removeInConfiguration(number-1)
		End If
	End If
	
End Sub

'Remove configuration nodes in XML based on position
Function removeInConfiguration(pos)
	
	'Else load XML FILE
	Set xmlDoc = loadXML()
	
	Set nodes = xmlDoc.SelectNodes("//Index")
	
	If pos < nodes.length Then
		'Remove and save
		nodes(pos).ParentNode.RemoveChild(nodes(pos))
		xmlDoc.Save(xmlpath)
	Else
		WScript.Echo "ID invalid"
	End If
	
End Function

'-------------------------------------------------------------------------------------------------------------
' END CONFIGURATION FUNCTIONS




' XML FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

Function loadXML()
	'Load XML FILE
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = "False"
	xmlDoc.Load(xmlpath)
	
	Set loadXML = xmlDoc
End Function


'Create the configuration file if the user wants it
Sub createConfigurationFile()
	WScript.Echo "Do you want to create the configration file ? (y/n)"
	
	'Confirm action
	If(askConfirmation()) Then
		WScript.Echo "Creating configuration file..."
		
		'Generate XML file with the root "Configuration"
		Set xmlDoc = _
		CreateObject("Microsoft.XMLDOM")  

		Set objRoot = _
		  xmlDoc.createElement("Configuration")  
		xmlDoc.appendChild objRoot  
	
		
		Set objIntro = _
		  xmlDoc.createProcessingInstruction _
		  ("xml","version='1.0'")  
		xmlDoc.insertBefore _
		  objIntro,xmlDoc.childNodes(0)  
		
		'Saving the XML file
		xmlDoc.Save xmlpath 
		
		WScript.Echo "Success!"
	Else
		WScript.Echo "No file has been created"
	End If
End Sub

'Check if configuration exists, if it does not, ask the user if he wants to create it
Function configurationExists()
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If(fso.FileExists(xmlpath)) Then
		configurationExists = True
	Else
		displayError("No configuration file exists")
		Call createConfigurationFile()
		configurationExists = False
	End If
End Function

'Load all indexes in the xml configuration and return dictionnary
Function getXMLIndexes()
	Set xmlDoc = loadXML()
	
	'Create lists
	Set activeDirectory = CreateObject("System.Collections.ArrayList")
	Set pronote = CreateObject("System.Collections.ArrayList")
	
	'Unique list in order to allow search functions to foreach only once
	Set uniqueActiveDirectory = CreateObject("System.Collections.ArrayList")
	
	'Foreach XML file
	For Each Index In xmlDoc.documentElement.SelectNodes("//Index")
	
		a = Index.getElementsByTagName("ActiveDirectory")(0).text
		'Add to the list the AD index
		activeDirectory.Add a
		pronote.Add Index.getElementsByTagName("Pronote")(0).text
		
		If Not uniqueActiveDirectory.Contains(a) Then
			uniqueActiveDirectory.Add a
		End If
	Next
	
	'Return dictionnary
	Set d= CreateObject("Scripting.Dictionary")
	d.Add "activeDirectory", activeDirectory
	d.Add "pronote", pronote
	d.Add "uniqueActiveDirectory", uniqueActiveDirectory
	
	Set getXMLIndexes = d
End Function


'-------------------------------------------------------------------------------------------------------------
' END XML FUNCTIONS



' UTILITIES FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

'Display error message
Sub displayError(message)
	WScript.Echo vbLf & "Error : " & message & vbLf
End Sub

'There is no max function in VBS
Function max(a, b)

	If a > b Then
		max = a
	Else
		max = b
	End If
	
End Function

'Handle spacing to display a table
Function getTableSpace(text, maxSize)
	spaceToAdd = ""
	
	offset = maxSize - Len(text) - 1
	
	If offset > 0 Then
		For i = 0 To offset Step 1
			spaceToAdd = spaceToAdd & " "
		Next
	End If
	
	getTableSpace = spaceToAdd
	
End Function

'Border to display a table
Function generateTextBorder(size)
	text = ""
	
	For i = 0 To size Step 1
		text = text & "-"
	Next
	
	generateTextBorder = text
	
End Function

'Ask the user for an input number, return -1 if not valid else the number
function askInputNumber()
	choice = WScript.StdIn.ReadLine
	
	If IsNumeric(choice) Then
		askInputNumber = CLng(choice)
	Else
		askInputNumber = -1
	End If
	
	'print space
	WScript.Echo ""
	
End function

'Ask confirmation, "y" for yes, everything else for no, return boolean
function askConfirmation()
	choice = WScript.StdIn.ReadLine
	
	If choice = "y" Then
		askConfirmation = True
	Else
		askConfirmation = False
	End If
	
	'print space
	WScript.Echo ""
	
End Function

'-------------------------------------------------------------------------------------------------------------
' END UTILITIES FUNCTIONS