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
Set objExcel = CreateObject("Excel.Application")

xmlpath = fso.BuildPath("C:\Users\vboissiere\Google Drive", "\" & "index.xml")
excelpath = "C:\Users\vboissiere\Google Drive\pronote script\elevesdetest2.xlsx"
myLdapPath = "dc=claudel,dc=lan"
userFriendlyOldDirectory = "Utilisateurs/Anciens/Anciens Eleves"

'Start foreach excel line with this number
excelLine = 2

excelLastNameCol = 1
excelFirstNameCol = 2
excelClassNameCol = 6
excelEmailCol = 5
excelDateCol = 8
excelNationalNumber = 7

'Difference in days between today's date and the date of "date de sortie". Positive integer
dayDiff = 90

'Do not change anything in ActiveDirectory
logModeOnly = True
'-------------------------------------------------------------------------------------------------------------
' END PARAMETERS


' MAIN PROGRAM
'-------------------------------------------------------------------------------------------------------------

WScript.Echo "SYNCING PROGRAM FOR PRONOTE -- Choose options"

'Needed global var
textLogWarning = ""
textLogCreated = ""
textLogUpdated = ""
textLogMoved = ""

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

	'Reset logs
	textLogWarning = ""
	textLogCreated = ""
	textLogUpdated = ""
	textLogMoved = ""

 	'Open excel document
	Set excelDoc = objExcel.Workbooks.Open(excelpath)
	currentLine = excelLine
	

	'Get all active Directory paths in the XML configuration file
	Set indexes = getXMLIndexes()

	
	'Add path for old Students directory, the pronote index is empty (thats how we can recognize it)
	indexes.Item("activeDirectory").Add userFriendlyOldDirectory
	indexes.Item("uniqueActiveDirectory").Add userFriendlyOldDirectory
	indexes.Item("pronote").Add ""
	
	
	WScript.Echo "Indexing Active Directory...."
	Set activeDirectoryGUID = getActiveDirectoryGUID(indexes.Item("uniqueActiveDirectory"))
	WScript.Echo "Done!"
	
	'Getting needed objects for the search
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection

	WScript.Echo "Updating students..."
	
	'Foreach students in excel
	Do Until objExcel.Cells(currentLine, excelClassNameCol).Value = ""
	
		'Get the student class in Excel
		studentCurrentClass = objExcel.Cells(currentLine,excelClassNameCol)
		studentDate = objExcel.Cells(currentLine,excelDateCol)
		
		'If the class exists in the xml configuration file, then search student
		If indexes.Item("pronote").Contains(studentCurrentClass) Then
			
			firstName = objExcel.Cells(currentLine,excelFirstNameCol)
			lastName = objExcel.Cells(currentLine,excelLastNameCol)
			
			'search student based on name
			Call searchStudent(objCommand, firstName & " " & lastName, indexes, studentCurrentClass, studentDate, currentLine)
		End If
		
		currentLine = currentLine + 1
	Loop
	
	WScript.Echo "Done!"
	
	If logModeOnly Then
		WScript.Echo vbLf & vbLf & "LOG MODE ONLY - NO CHANGE IN ACTIVE DIRECTORY"
		WScript.Echo "------------------------------------------------" & vbLf
	End If
	
	WScript.Echo vbLf & "STUDENTS UPDATED : "
	WScript.Echo textLogUpdated
	
	WScript.Echo vbLf & "STUDENTS MOVED : "
	WScript.Echo textLogMoved
	
	WScript.Echo vbLf & "STUDENTS CREATED : "
	WScript.Echo textLogCreated
	
	WScript.Echo vbLf & "LIST OF WARNINGS : "
	WScript.Echo textLogWarning
	
	'Close connection to AD
	objConnection.Close
	
	objExcel.Quit
	
End Sub

'Create student based on Excel data
Sub createStudent(currentLine, indexes, studentCurrentClass)

	firstName = objExcel.Cells(currentLine,excelFirstNameCol).Text
	lastName = objExcel.Cells(currentLine,excelLastNameCol).Text
	className = objExcel.Cells(currentLine,excelClassNameCol).Text
	email = objExcel.Cells(currentLine,excelEmailCol).Text
	nationalNumber = objExcel.Cells(currentLine,excelNationalNumber).Text
	
	'Get the right activeDirectory position in index
	activeDirectoryPos = indexes.Item("pronote").IndexOf(studentCurrentClass, 0)
	
	'Get active directory friendly path
	friendlyPath = indexes.Item("activeDirectory").Item(activeDirectoryPos)
	
	'Get login for new user
	login = getLogin(firstName, lastName)
	
	If Not logModeOnly Then
		
		Set userObj = getActiveOUDirectory(friendlyPath)
		
		If Not userObj Is Nothing Then
			Set objUser = userObj.Create("User", "cn= "& firstName & " " & lastName)
			objUser.Put "firstName", firstName
			objuser.Put "lastName", lastName
			objUser.put "mail", email
			objUser.put "description", className
			objUser.Put "displayName", firstName & " " & lastName & " " & className
			objUser.Put "userPrincipalName", login
			objUser.Put "sAMAccountName", login
			objUser.Put "physicalDeliveryOfficeName", nationalNumber
			objUser.SetInfo
		End If
		
	End If

	textLogCreated = textLogCreated & vbLf & firstName & " " & lastName & " in " & friendlyPath & ". Login : " & login
	
End Sub

'Move student to the new right path
Sub moveStudent(student, friendlyPath)

	rawPath = getLdapPath(friendlyPath)
	
	'Move student
	If Not logModeOnly Then
		objOU.MoveHere _
    	rawPath, vbNullString
    End If
    
    'log
	textLogMoved = textLogMoved & vbLf & student.Firstname & " " & student.lastName & " moved to " & friendlyPath
	
End Sub

'Trigger actions based on the Excel data and the position of the student and its category (active vs old)
Sub studentExists(studentCurrentClass, studentName, indexes, posFound, IsOld, student, currentLine)

	'WScript.Echo "Student Exist... (actions simulated)"
	
	posShouldBeIn = indexes.Item("pronote").IndexOf(studentCurrentClass, 0)
	
	'Check if found in old or active. Last is the old directory
	If posFound = indexes.Item("uniqueActiveDirectory").Count - 1 Then
		
		'Move student if should be in active path and is in old path
		If Not IsOld Then
			Call moveStudent(student, indexes.Item("activeDirectory").Item(posShouldBeIn))
		End If
		
		Call updateStudent(student, currentLine)
		
	Else
		'Active path
		
		If IsOld Then
			Call moveStudent(student, indexes.Item("uniqueActiveDirectory").Item(indexes.Item("uniqueActiveDirectory").Count - 1))
		
		Else
			'Good category
			
			pronoteADIndex = indexes.Item("uniqueActiveDirectory").Item(posFound)
		
			'Wrong section
			If indexes.Item("activeDirectory").Item(posShouldBeIn) <> pronoteAdIndex Then
				Call moveStudent(student, indexes.Item("activeDirectory").Item(posShouldBeIn))
			End If
		End If
		
		Call updateStudent(student, currentLine)
		
	End If

	
End Sub

'Update student data and display it in the console screen
Sub updateStudent(student, currentLine)

	firstName = objExcel.Cells(currentLine,excelFirstNameCol).Text
	lastName = objExcel.Cells(currentLine,excelLastNameCol).Text
	className = objExcel.Cells(currentLine,excelClassNameCol).Text
	email = objExcel.Cells(currentLine,excelEmailCol).Text
	nationalNumber = objExcel.Cells(currentLine,excelNationalNumber).Text
	
	textLog = ""
	
	If firstName <> student.FirstName Then
		textLog = textLog & " firstName: " & firstName
	End If
	
	If lastName <> student.LastName Then
		textLog = textLog & " lastName: " & lastName
	End If
	
	If className <> student.Description Then
		textLog = textLog & " ClassName: " & className
	End If
	
	'If email not set in pronote, display only a warning
	If email = "#N/A"  Then
		textLogWarning = textLogWarning & vbLf &  "WARNING : " & student.cn & " has no email set on Pronote, active directory email untouched"
	Else If email <> student.EmailAddress Then
		textLog = textLog & " email: " & email
		End If
	End If
	
	'If nationalNumber <> student.physicalDeliveryOfficeName Then
	'	textLog = textLog & " current National Number: " & nationalNumber
	'End If
	
	'save modification only if log mode Only is false
	If Not logModeOnly Then
		student.Put "firstName", firstName
		student.Put "lastName", lastName
		student.Put "description", className
		student.Put "physicalDeliveryOfficeName", nationalNumber
		
		'No warning on the email
		If email <> "#N/A" Then
			student.Put "mail", email
		End If
		
		student.Put "displayName", firstName & " " & lastName & " " & className
		student.setInfo
	End If
	
	If textLog <> "" Then
		textLogUpdated = textLogUpdated & vbLf & student.cn & ". " & textLog
	End If
	
End Sub

'Search student. Can be used many times
Sub searchStudent(objCommand, studentName, indexes, studentCurrentClass, studentDate, currentLine)

	Set uniqueActiveDirectory = indexes.Item("uniqueActiveDirectory")

	'Search students in the given configuration indexes (AD paths)
	For i = 0 To uniqueActiveDirectory.Count - 1 Step 1
		
		'Query for the search
		objCommand.CommandText = _
		    "<" & getLdapPath(uniqueActiveDirectory.Item(i)) & _
		     ">;(&(objectCategory=person)(objectClass=user)(displayName=" & studentName & "*));cn;subtree"
		  
		Set objRecordSet = objCommand.Execute
		 
		numberOfMatch = objRecordset.RecordCount
		
		'if one student found, stop search and execute subRoutine
		If numberOfMatch = 1 Then
		
			studentCN = objRecordSet.Fields("cn").Value
			
			rawPath = "LDAP://cn=" & studentCN & "," & getSmallLdapPath(uniqueActiveDirectory.Item(i))
			
    		Set student = getActiveOUDDirectoryFromRaw(rawPath, rawPath)
    		
    		If Not student Is Nothing Then
			
		    	Call studentExists(studentCurrentClass, studentName, indexes, i, studentDate <> "" And DateDiff("d",Now, studentDate) < dayDiff, student, currentLine)
		    	
		    End If
		    
		    objRecordSet.Close
		    Exit Sub
		Else If numberOfMatch > 1 Then
			textLogWarning = textLogWarning & vbLf &  "WARNING : More than one match for " & studentName & " (User ignored)"
			Exit Sub
			End If
		End If
	Next

	'No student found, create student
	Call createStudent(currentLine, indexes,studentCurrentClass)
	objRecordSet.Close
End Sub


'-------------------------------------------------------------------------------------------------------------
' END SYNCING FUNCTIONS




' ACTIVE DIRECTORY FUNCTIONS
'-------------------------------------------------------------------------------------------------------------

'Get login based on student name
Function getLogin(firstName, lastName)

	'Regex pattern to get a nice login
	Set objReg = CreateObject("VBScript.RegExp")
	objReg.Global = True
	objReg.Pattern = "[^A-Za-z]"
	

	login = Mid(getURLLikeString(firstName, objReg),1,3) & Mid(getURLLikeString(lastName, objReg),1,3)

	'Declare objects for the search in AD
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	
	'if has fewer than 6 characters, add random ones
	Do While Len(login) < 6
  		login = login & strRandom()
	Loop
	
	nbText = ""
	nb = 1
	
	'Verify the uniqueness of the login
	Do 
	
	'search query
	objCommand.CommandText = _
		    "<LDAP://" & myLdapPath & _
		     ">;(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & login & nb & "*));cn;subtree"
		  
	Set objRecordSet = objCommand.Execute
			 
	'get number of students with this username
	numberOfMatch = objRecordset.RecordCount
	
	objRecordSet.Close
	
	'next iteration
	nb = nb + 1
	nbText = CStr(nb)
	
	Loop While numberOfMatch > 0
	
	'close connection for the search
	objConnection.Close
	
	If nb > 2 Then
		getLogin = login
	Else
		getLogin = login & nbText
	End If
	
End Function

'Everything in the LDAP except the beginning
Function getSmallLdapPath(path)

	'The path should be written as something like "Utilisateurs/Eleves/Eleves du CM2
	OUarray = Split(path,"/")
	
	'Get all OU, with the right slash
	ouPath = ""
	
	'concatenate in reverse order for nicer software (more user friendly)
	For Each x In OUarray
		ouPath = "ou=" & x & "," & ouPath
	Next
	
	getSmallLdapPath = ouPath & myLdapPath

End Function

'get ldap path based on the path written in the configuration XML file
Function getLdapPath(path)
	
	getLdapPath = "LDAP://" & getSmallLdapPath(path)

End Function

'Get the folder Object to a Active Directory folder. Return Nothing if path is false, errPath is the path showed in the error message
Function getActiveOUDDirectoryFromRaw(rawPath, errPath)
	
	If rawPath <> "" Then
	
		'Check if OU exists, return the OU object if that is the case, else return null
		On Error Resume Next
		Dim OUobject
		Set OUobject = GetObject(rawPath)
		
		'Display corresponding errors, return null if error(s) else return OU
		Select Case Err.Number
		Case 0
		    Set getActiveOUDDirectoryFromRaw = OUobject
		Case &h80072030
		    displayError("OU doesn't exist" & vbLf & vbLf & "Full path : " & errPath & vbLf)
		    Set getActiveOUDDirectoryFromRaw = Nothing
		Case Else
		    displayError("Adding OU failed because OU not valid. Error code : "& Err.Number & vbLf & vbLf & "Full path : " & errPath & vbLf)
		    Set getActiveOUDDirectoryFromRaw = Nothing
		End Select
	Else
		displayError("Wrong format path to active Directory")
		Set getActiveOUDDirectoryFromRaw = Nothing
	End If
	
	
End Function

'Get the path from a friendlyPath like Utilisateurs/Eleves/Eleves du College
Function getActiveOUDirectory(friendlyPath)
	
	Set getActiveOUDirectory = getActiveOUDDirectoryFromRaw(getLdapPath(friendlyPath), friendlyPath)

End Function

'Check if data given to add the the xml file is valid
Function validateIndex(pronote, activeDirectory)
	
	If Len(pronote) <= 1 or Len(activeDirectory) <= 1 Then
		displayError("Lenght of Pronote index and active directory index should be both greather than 1")
		validateIndex = False
	Else
		'Check if pronote is unique (not already added)
		Set indexes = getXMLIndexes()
		validateIndex = indexes("pronote").IndexOf(pronote, 0) = -1
		
		If Not validateIndex Then
			displayError("L'index de pronote doit être unique.")
		End If
	End If

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
	WScript.Echo vbLf & "ERROR : " & message & vbLf
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
	
	askConfirmation = choice = "y"
	
	'print space
	WScript.Echo ""
	
End Function

'Random letter for user login
Function strRandom()
	Randomize
    
    'the random letter
    strRandom = Mid("abcdefghijklmnopqrstuvwxyz",Int((25) * Rnd + 1) ,1)
  
End Function

'Return the text with the objReg (remove all char that match Regex code)
Function getURLLikeString(txt, objReg)
	txt = LCase(txt)
	txt = Replace(txt,"é","e")
	txt = Replace(txt,"É","e")
	txt = Replace(txt,"ï","i")
	getURLLikeString = objReg.Replace(txt,"")
End Function


'-------------------------------------------------------------------------------------------------------------
' END UTILITIES FUNCTIONS