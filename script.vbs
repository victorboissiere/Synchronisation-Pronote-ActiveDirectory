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

xmlpath = fso.BuildPath("C:\users\vboissiere\Google Drive\", "index.xml")
excelpath = "C:\users\vboissiere\Google Drive\pronote script\eleve juin 2015.xlsx"
myLdapPath = "DC=claudel,DC=lan"

'Paths in order to create a new student
profilepath = "\\claudel.lan\partage\profils$\" '(+login added later)
homeDirectory = "\\claudel.lan\partage\utilisateurs$\"
homeDrive = "P"

'Start foreach excel line with this number
excelLine = 2

excelLastNameCol = 1
excelFirstNameCol = 2
excelClassNameCol = 6
excelEmailCol = 5
excelDateCol = 8
excelNationalNumber = 7

'Difference in days between today's date and the date of "date de sortie". Positive integer
dayDiff = 0

'WARNING : be careful to respect Active Directory requirements or students will have no password (but account will be disabled)
defaultPassword = "Passw0rd"

'OLD PATH
oldGroupPath = "Utilisateurs/Anciens/Groupe des anciens Eleves"
userFriendlyOldDirectory = "Utilisateurs/Anciens/Anciens Eleves"
Dim oldSubDirectories(2)
oldSubDirectories(0)= "Annee2013"
oldSubDirectories(1)= "Annee2014"
oldSubDirectories(2)= "Annee2015"
'-------------------------------------------------------------------------------------------------------------
' END PARAMETERS


' MAIN PROGRAM
'-------------------------------------------------------------------------------------------------------------

WScript.Echo "SYNCING PROGRAM FOR PRONOTE -- Choose options"

'LOGS

'Variables for the tables
Set textLogMoved = CreateObject("System.Collections.ArrayList") 
Set textLogUpdated = CreateObject("System.Collections.ArrayList")
Set textLogCreated = CreateObject("System.Collections.ArrayList")

'OTHERS
textLogWarning = ""
textLogError = ""
logModeOnly = True

'Main loop of the program
Do While True

	'Display the menu
	WScript.Echo vbLf & "1. Sync" & vbLf & "2. Reset Password" & vbLf & "3. Display index" & vbLf & "4. Add index" & vbLf & "5. Remove index" & vbLf & "0. Exit"
	
	choice = askInputNumber()
	
	'Check option and trigger corresponding functions
	If choice >= 0 Then
		
		'Each function verify that the configuration exists in order to run
		Select Case True
			Case choice = 0
				WScript.Quit
			'Check configuration exists else ask user if he wants to create it
			Case Not configurationExists()
			Case choice = 1
				sync()
			Case choice = 2
				resetPassword()
			Case choice = 3
				'Display table
				displayConfiguration()
			Case choice = 4
				'Ask user data and add it if valid to the XML configuration file
				writeConfiguration()
			Case choice = 5
				'Ask user ID of index and remove it from the XML configuration file
				removeConfiguration()
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
	textLogError = ""
	
	Call generateHeaders()
	
	WScript.Echo "Do you want to run the script in production mode ? (y/n)"
	
	logModeOnly = Not askConfirmation()
	

 	'Open excel document
	Set excelDoc = objExcel.Workbooks.Open(excelpath)
	currentLine = excelLine
	

	'Get all active Directory paths in the XML configuration file
	Set indexes = getXMLIndexes()

	
	'Add path for old Students directory, the pronote index is empty (thats how we can recognize it)
	indexes.Item("activeDirectory").Add userFriendlyOldDirectory
	indexes.Item("uniqueActiveDirectory").Add userFriendlyOldDirectory
	indexes.Item("pronote").Add ""
	
	'Adding all anciens directory old paths if they are valid
	For Each oldPath In oldSubDirectories
		thisPath = userFriendlyOldDirectory & "/" & oldPath
		
		Set checkOU = getActiveOUDirectory(thisPath)
		If Not checkOU Is Nothing And thisPath <> userFriendlyOldDirectory Then
			indexes.Item("activeDirectory").Add thisPath
			indexes.Item("uniqueActiveDirectory").Add thisPath
			indexes.Item("pronote").Add ""
		Else 
			WScript.Echo "WARNING : the old path " & thisPath & " is not valid or already exists"
		End If
	Next
	
	
	WScript.Echo "Indexing Active Directory...."
	Set activeDirectoryGUID = getActiveDirectoryGUID(indexes.Item("uniqueActiveDirectory"), indexes)
	WScript.Echo "Done!"
	
	'Getting needed objects for the search
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection

	WScript.Echo "Updating students..."
	
	'Foreach students in excel
	Do Until objExcel.Cells(currentLine, excelFirstNameCol).Value = ""
	
		'Get the student class in Excel
		studentCurrentClass = objExcel.Cells(currentLine,excelClassNameCol)
		studentDate = objExcel.Cells(currentLine,excelDateCol)
		
		'If the class exists in the xml configuration file or is empty, then search student
		If indexes.Item("pronote").Contains(studentCurrentClass) Or studentCurrentClass = "" Then
		
			
			firstName = objExcel.Cells(currentLine,excelFirstNameCol)
			lastName = objExcel.Cells(currentLine,excelLastNameCol)
			
			'search student based on name
			Call searchStudent(objCommand, firstName & " " & lastName, indexes, studentCurrentClass, studentDate, currentLine, activeDirectoryGUID)
		
		End If
		currentLine = currentLine + 1
	Loop
	
	WScript.Echo "Done!"
	
	If logModeOnly Then
		WScript.Echo vbLf & vbLf & "LOG MODE ONLY - NO CHANGE IN ACTIVE DIRECTORY"
		WScript.Echo "------------------------------------------------" & vbLf
	End If
	
	WScript.Echo vbLf & "STUDENTS UPDATED : " & vbLf
	Call displayTable(textLogUpdated)
	
	WScript.Echo vbLf & "STUDENTS MOVED : " & vbLf
	'Add headers for the table
	Call displayTable(textLogMoved)
	
	WScript.Echo vbLf & "STUDENTS CREATED : " & vbLf
	Call displayTable(textLogCreated)
	
	WScript.Echo vbLf & "LIST OF WARNINGS : "
	WScript.Echo textLogWarning
	
	WScript.Echo vbLf & "LIST OF ERROR : "
	WScript.Echo textLogError
	
	WScript.Echo vbLf & "LIST OF PEOPLE NOT FOUND IN PRONOTE : " & vbLf
	
	'All students that match the class but that are not in ProNote
	For Each cn In activeDirectoryGUID
		begin = InStr(cn, "=") + 1
		last = InStr(cn, ",")
		WScript.Echo Mid(cn, begin, last-begin)
	Next
	
	'Close connection to AD
	objConnection.Close
	
	objExcel.Quit
	
End Sub

'Search student. Can be used many times
Sub searchStudent(objCommand, studentName, indexes, studentCurrentClass, studentDate, currentLine, activeDirectoryGUID)

	Set uniqueActiveDirectory = indexes.Item("uniqueActiveDirectory")

	'Search students in the given configuration indexes (AD paths)
	For i = 0 To uniqueActiveDirectory.Count - 1 Step 1
		
		'Query for the search
		objCommand.CommandText = _
		    "<" & getLdapPath(uniqueActiveDirectory.Item(i)) & _
		     ">;(&(objectCategory=person)(objectClass=user)(displayName=" & studentName & "*));cn;onelevel"
		  
		Set objRecordSet = objCommand.Execute
		 
		numberOfMatch = objRecordset.RecordCount
		
		'if one student found, stop search and execute subRoutine
		If numberOfMatch = 1 Then
		
			studentCN = objRecordSet.Fields("cn").Value
			
			rawPath = "LDAP://CN=" & studentCN & "," & getSmallLdapPath(uniqueActiveDirectory.Item(i))
			
    		Set student = getActiveOUDDirectoryFromRaw(rawPath, rawPath)
    		
    		If Not student Is Nothing Then
    			'Remove from index
    			activeDirectoryGUID.remove(LCase(student.ADspath))
		    	Call studentExists(studentCurrentClass, studentName, indexes, i, studentDate <> "" And DateDiff("d",Now, studentDate) < dayDiff, student, currentLine)
		    	'WScript.Echo student.Cn
		    End If
		    
		    objRecordSet.Close
		    Exit Sub
		Else If numberOfMatch > 1 Then
			textLogWarning = textLogWarning & vbLf &  "WARNING : More than one match for " & studentName & " (User ignored)"
			Exit Sub
			End If
		End If
	Next

	'No student found, create student only if class is not empty
	If studentCurrentClass <> "" Then
		Call createStudent(currentLine, indexes,studentCurrentClass)
	End If
	objRecordSet.Close
End Sub

'Trigger actions based on the Excel data and the position of the student and its category (active vs old)
Sub studentExists(studentCurrentClass, studentName, indexes, posFound, IsOld, student, currentLine)

	'Student exist but class is null, move it the old path
	If studentCurrentClass = "" Then
		'move only if not already found in old path
		If posFound < indexes.Item("uniqueActiveDirectory").Count - UBound(oldSubDirectories) - 2 Then
			
			'TODO : CHECK IF IT IS WORKING FOR THURSDAY
			Set group = student.Groups
			Dim lastGroup
			nbGroupPath = 0
			
			For Each g In group
				Set lastGroup = g
				nbGroup = nbGroup + 1
			Next
			
			If nbGroup <= 1 Then
				lastGroupPath = LCase(lastGroup.ADsPath)
			End If
			
			If nbGroup <= 1 Then
			
				groupPath = getGroupPath(oldGroupPath)

				Set objGroup = getActiveOUDDirectoryFromRaw(groupPath, groupPath)
				
				If Not objGroup Is Nothing And Not logModeOnly Then
				
					'delete old Group if exists
					If nbGroup = 1 Then
	        			lastGroup.remove(student.ADsPath)
	        		End If
	        	
					objGroup.add(student.ADsPath)
				End If
			Else If nbGroup > 1 Then
				textLogWarning = textLogWarning & vbLf &  "WARNING : " & student.cn & " has more than 1 group. No group changed."
				End If
			End If
			
			
			Call moveStudent(student, indexes.Item("uniqueActiveDirectory").Item(indexes.Item("uniqueActiveDirectory").Count - 1))
		End If
		Exit Sub
	End If
		
		
	posShouldBeIn = indexes.Item("pronote").IndexOf(studentCurrentClass, 0)
	
	'Check if found in old or active. Last is the old directory
	If posFound >= indexes.Item("uniqueActiveDirectory").Count - UBound(oldSubDirectories) - 2 Then
	
		
		'Move student if should be in active path and is in old path
		If Not IsOld Then
			Call updateStudent(student, currentLine, indexes, studentCurrentClass)
			Call moveStudent(student, indexes.Item("activeDirectory").Item(posShouldBeIn))
			Exit Sub
		End If
		
		Call updateStudent(student, currentLine, indexes, studentCurrentClass)
		
	Else
		'Active path
		
		If IsOld Then
			Call updateStudent(student, currentLine, indexes, studentCurrentClass)
			Call moveStudent(student, indexes.Item("uniqueActiveDirectory").Item(indexes.Item("uniqueActiveDirectory").Count - 1))
			Exit Sub
		Else
			'Good category
			
			pronoteADIndex = indexes.Item("uniqueActiveDirectory").Item(posFound)
		
			'Wrong section
			If indexes.Item("activeDirectory").Item(posShouldBeIn) <> pronoteAdIndex Then
				Call updateStudent(student, currentLine, indexes, studentCurrentClass)
				Call moveStudent(student, indexes.Item("activeDirectory").Item(posShouldBeIn))
				Exit Sub
			End If
		End If
		
		Call updateStudent(student, currentLine, indexes, studentCurrentClass)
		
	End If

	
End Sub

'Create student based on Excel data
Sub createStudent(currentLine, indexes, studentCurrentClass)

	firstName = objExcel.Cells(currentLine,excelFirstNameCol).Text
	lastName = objExcel.Cells(currentLine,excelLastNameCol).Text
	className = objExcel.Cells(currentLine,excelClassNameCol).Text
	email = objExcel.Cells(currentLine,excelEmailCol).Text
	nationalNumber = objExcel.Cells(currentLine,excelNationalNumber).Text
	
	If InStr(email, "claudel.org") = 0 Then
		textLogWarning = textLogWarning & vbLf &  "WARNING : " & firstName & " " & lastName & _
				" has the email " & email & " which is not claudel.org"
	End If
	
	'Get the right activeDirectory position in index
	activeDirectoryPos = indexes.Item("pronote").IndexOf(studentCurrentClass, 0)
	
	'Get active directory friendly path
	friendlyPath = indexes.Item("activeDirectory").Item(activeDirectoryPos)
	
	'Get login for new user
	login = getLogin(firstName, lastName)
	
	If Not logModeOnly Then
		
		Set userObj = getActiveOUDirectory(friendlyPath)
		
		If Not userObj Is Nothing Then
			Set objUser = userObj.Create("User", "CN="& firstName & " " & lastName)
			
			'Account properties
			objUser.firstName = firstName
			objUser.lastName = lastName
			objUser.cn = firstName & " " & lastName
			If InStr(email, "claudel.org") <> 0 Then
				objUser.mail = email				
			End If
			objUser.description = className
			objUser.displayName = firstName & " " & lastName & " " & className
			objUser.userPrincipalName = login
			objUser.sAMAccountName = login
			objUser.physicalDeliveryOfficeName = nationalNumber
			objUser.profilePath = profilepath & login
			objUser.homeDrive = homeDrive
			objUser.homeDirectory = homedirectory & login
			objUser.SetInfo
			
			'VBS equivalent try catch, error if password does not meet active directory requirements
			On Error Resume Next
			Err.Clear
			
			objUser.setPassword(defaultPassword)
			
			If Err.Number <> 0 Then
				textLogError = textLogError & vbLf &  "ERROR : Password does not match active Directory requirements." & _ 
				vbLf & firstName & " " & lastName & " student created with no password but account is disabled."
			End If
			
			objUser.AccountDisabled=False
			objUser.pwdLastSet=0
			objUser.SetInfo

			
			'Accounts settings
			
			groupPath = getGroupPath(indexes.Item("group").Item(activeDirectoryPos))
			
			
			Set objGroup = getActiveOUDDirectoryFromRaw(groupPath, groupPath)
			
			If Not objGroup Is Nothing Then
				objGroup.add(objUser.ADsPath)
			End If
			
			
		End If
		
	End If
	
	Set column = CreateObject("System.Collections.ArrayList") 
	column.Add firstName & " " & lastName
	column.Add friendlyPath
	column.Add login
	textLogCreated.Add column
	
End Sub

'Move student to the new right path
Sub moveStudent(student, friendlyPath)
	
	Set ou = getActiveOUDirectory(friendlyPath)
	
	If Not ou Is Nothing Then
	
		'log
		Set movedTo = CreateObject("System.Collections.ArrayList")
		movedTo.Add student.Firstname & " " & student.lastName
		movedTo.Add getFriendlyPathFromOU(student.ADsPath)
		movedTo.Add friendlyPath
		
		textLogMoved.Add movedTo
		
		'Move student
		If Not logModeOnly Then
			WScript.Echo "move here " & student.ADsPath
			ou.MoveHere student.ADsPath, vbNullString
	    End If
		
	End If
	
End Sub

'Update student data and display it in the console screen
Sub updateStudent(student, currentLine, indexes, studentCurrentClass)

	firstName = objExcel.Cells(currentLine,excelFirstNameCol).Text
	lastName = objExcel.Cells(currentLine,excelLastNameCol).Text
	className = objExcel.Cells(currentLine,excelClassNameCol).Text
	email = objExcel.Cells(currentLine,excelEmailCol).Text
	nationalNumber = objExcel.Cells(currentLine,excelNationalNumber).Text
	
	pos = indexes.Item("pronote").IndexOf(studentCurrentClass, 0)
	shouldBeInGroupPath = LCase(getGroupPath(indexes.Item("group").Item(pos)))
	
	Set column = CreateObject("System.Collections.ArrayList") 
	column.Add ""'0 student
	column.Add ""'1 firstName
	column.Add ""'2 lastName
	column.Add ""'3 email
	column.Add ""'4 class
	column.Add ""'5 group
	
	nbModif = 0
	
	
	'foreach group. If equal than 1, compare, if 0 or more than 1, warning
	Set group = student.Groups
	Dim lastGroup
	nbGroupPath = 0
	
	For Each g In group
		Set lastGroup = g
		nbGroup = nbGroup + 1
	Next
	
	If nbGroup <= 1 Then
		lastGroupPath = LCase(lastGroup.ADsPath)
	End If
	
	textLog = ""
	
	If nbGroup <= 1 And StrComp(shouldBeInGroupPath, lastGroupPath, 1) <> 0 Then
		column.Item(5) = shouldBeInGroupPath
		nbModif = nbModif + 1
	Else If nbGroup > 1 Then
			textLogWarning = textLogWarning & vbLf &  "WARNING : " & student.cn & " has more than 1 group. No group changed."
		End If
	End If
	
	
	
	If firstName <> student.FirstName Then
		column.Item(1) = firstName
		nbModif = nbModif + 1
	End If
	
	If lastName <> student.LastName Then
		column.Item(2) = lastName
		nbModif = nbModif + 1
	End If
	
	If className <> student.Description Then
		column.Item(4) = className
		nbModif = nbModif + 1
	End If
	
	

	If email = "#N/A" Then
		textLogWarning = textLogWarning & vbLf &  "WARNING : " & student.cn & " has no email set on Pronote, active directory email untouched"		
	Else If email <> student.mail Then
		If InStr(email, "claudel.org") <> 0 Then
			column.Item(3) = email
			nbModif = nbModif + 1
		Else
			textLogWarning = textLogWarning & vbLf &  "WARNING : " & student.cn & " has the email " & email & " which is not claudel.org"		
		End If
		End If
	End If

	
	'If nationalNumber <> student.physicalDeliveryOfficeName Then
	'	textLog = textLog & " current National Number: " & nationalNumber
	'End If
	
	'save modification only if log mode Only is false
	If Not logModeOnly Then
		student.firstName = firstName
		student.lastName = lastName
		student.description = className
		student.physicalDeliveryOfficeName = nationalNumber
		
		'No warning on the email
		If email <> "#N/A" And InStr(email, "claudel.org") <> 0 Then
			student.mail = email
		End If
		
		'Change group
		If nbGroup <= 1 And StrComp(shouldBeInGroupPath, lastGroupPath, 1) <> 0 Then
		
			groupPath = getGroupPath(indexes.Item("group").Item(pos))
			Set objGroup = getActiveOUDDirectoryFromRaw(groupPath, groupPath)
			
			If Not objGroup Is Nothing Then
			
				'delete old Group if exists
				If nbGroup = 1 Then
        			lastGroup.remove(student.ADsPath)
        		End If
        	
				objGroup.add(student.ADsPath)
			End If
		End If
		
		student.displayName = firstName & " " & lastName & " " & className
		student.setInfo
	End If
	
	If nbModif > 0 Then
		column.Item(0) = student.firstName & " " & student.lastName
		textLogUpdated.Add column
	End If
	
End Sub


'-------------------------------------------------------------------------------------------------------------
' END SYNCING FUNCTIONS


' RESET PASSWORD FUNCTIONS
'-------------------------------------------------------------------------------------------------------------


'Reset password in an active Directory Path for student of the rightClass
Sub resetPasswordIn(activeDirectoryPath, password, studentClass, askReset)

	WScript.Echo vbLf & vbLf & "The student of the class " & studentClass & " will be asked to change their passwords"

	WScript.Echo vbLf & "LIST OF USERS UPDATED WITH NEW PASSWORD : " & vbLf	
	
	Set ou = getActiveOUDirectory(activeDirectorypath)
	
	'Verify OU exists
	If Not ou Is Nothing Then
			'For each student check if match the class
			For Each student in ou
				If student.Description = studentClass Then
					
					'VBS equivalent of TRY/CATCH
					On Error Resume Next
					Err.Clear
					If askReset Then
						student.pwdLastSet=0
					End If
					student.setPassword(password)
					student.setInfo
					
					'Check if password is valid, else stop and exit
					If Err.Number = 0 Then
						WScript.Echo student.cn & " has now the password " & password
					Else
						displayError("Password does not match active Directory requirements" & vbLf & "Aborted!")
						Exit Sub
					End If
					
				End If
			Next
	Else
		displayError("OU path not valid")
	End If

End Sub

'Ask index ID in order to reset password based on class
Sub resetPassword()
	WScript.Echo "ID de la classe pour modifier le mot de passe : "
	
	number = askInputNumber()
	
	If number <> -1 And number > 0 Then
		Set indexes = getXMLIndexes()
		If number <= indexes.Item("pronote").Count Then
		
			number = number -1 'Modify for the index
		
			'Data
			theClass = indexes.Item("pronote").Item(number)
			activeDirectoryPath = indexes.Item("activeDirectory").Item(number)
			studentClass = indexes.Item("pronote").Item(number)
			
			'ask confirmation
			WScript.Echo "Are you sure you want to reset the password of the " & _
					theClass & " class corresponding to the path " & _
					activeDirectoryPath & " in Active Directory ? (y/n)"
					
			confirmation = askConfirmation()
			
			'ask for password and trigger subMenu to reset password
			If confirmation Then
				WScript.Echo "Type the password you want to set for the " & theClass & " class"
				password = WScript.StdIn.ReadLine
				If password <> "" Then
					WScript.Echo vbLf & "Do you want to ask the class " & theClass & " to change the default password when they log in ? (y/n)"
					Call resetPasswordIn(activeDirectoryPath, password, studentClass, askConfirmation())
				Else
					displayError("Empty password!")
				End If
			Else
				WScript.Echo "Aborted!"
			End If
		Else
			displayError("The index does not exists")
		End If
	Else
		displayError("This is not a number or the ID cannot be 0 (start at 1)")
	End if

End Sub

'-------------------------------------------------------------------------------------------------------------
' END RESET PASSWORD FUNCTIONS


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
	nb = 0
	
	'Verify the uniqueness of the login
	Do 
		nb = nb + 1
		'search query
		objCommand.CommandText = _
			    "<LDAP://" & myLdapPath & _
			     ">;(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & login & nbText & "*));cn;subtree"
			  
		Set objRecordSet = objCommand.Execute
				 
		'get number of students with this username
		numberOfMatch = objRecordset.RecordCount
		
		objRecordSet.Close
		
		'next iteration
		nbText = CStr(nb)
	
	Loop While numberOfMatch > 0
	
	'close connection for the search
	objConnection.Close
	
	If nb <= 1 Then
		getLogin = login
	Else
		getLogin = login & nbText
	End If
	
End Function

'Everything in the LDAP except the beginning
Function getSmallLdapPath(friendlyPath)

	'The path should be written as something like "Utilisateurs/Eleves/Eleves du CM2
	OUarray = Split(friendlyPath,"/")
	
	'Get all OU, with the right slash
	ouPath = ""
	
	'concatenate in reverse order for nicer software (more user friendly)
	For Each x In OUarray
		ouPath = "OU=" & x & "," & ouPath
	Next
	
	getSmallLdapPath = ouPath & myLdapPath

End Function

'get ldap path based on the path written in the configuration XML file
Function getLdapPath(friendlyPath)
	
	getLdapPath = "LDAP://" & getSmallLdapPath(friendlyPath)

End Function

'Get group path (add CN in the front)
Function getGroupPath(path)

	OUarray = Split(path, "/")
	
	ouPath = ""
	
	For i = 0 To UBound(OUarray) - 1 Step 1
		ouPath = "OU=" & OUarray(i) & "," & ouPath
	Next
	
	getGroupPath = "LDAP://CN=" & OUarray(UBound(OUarray)) & "," & ouPath  & myLdapPath
	
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
Function getActiveDirectoryGUID(paths, indexes)
	
	'Create the list
	Set activeDirectoryGUID = CreateObject("System.Collections.ArrayList")
	
	'Foreach all AD paths in the XML configuration file
	For i = 0 To paths.Count - 2 - UBound(oldSubDirectories) Step 1
		
		'Get OU
		Set OUs = getActiveOUDirectory(paths.Item(i))
		
		
		'Check if valid and add unique GUID to the list
		If Not OUs Is Nothing Then
			For Each user in OUs
				If indexes.Item("pronote").IndexOf(user.Description, 0) <> -1 Then
    				activeDirectoryGUID.Add LCase(user.ADsPath)
    			End If
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
	
	'Get nodes
	Set nodes = xmlDoc.documentElement.SelectNodes("//Index")
	
	'If none, display message and exit function
	If nodes.length = 0 Then
		WScript.Echo "There is no index at the time"
		Exit Sub
	End If
	
	'Create first row
	Set rows = CreateObject("System.Collections.ArrayList")
	Set headers = CreateObject("System.Collections.ArrayList")
	headers.Add "Class"
	headers.Add "Active Directory Path"
	headers.Add "Active Directory Group Path"
	rows.Add headers
	
	'Foreach XML file
	For Each Index In nodes
	
		Set column = CreateObject("System.Collections.ArrayList")
	
		'Get tag name in the XML
		column.Add Index.getElementsByTagName("Pronote")(0).text
		column.Add Index.getElementsByTagName("ActiveDirectory")(0).text
		column.Add Index.getElementsByTagName("Group")(0).text
		
		rows.Add column
	Next
	
	displayTable(rows)

	
End Sub

'Write into the current configuration of the xml file that contains the settings
Sub writeConfiguration()

	'Ask the user the right configuration
	WScript.Echo "Name of the class in Pronote (column " & excelClassNameCol & ") : "
	pronote = WScript.StdIn.ReadLine
	
	WScript.Echo vbLf & "Corresponding path in Active Directory : "
	activeDirectory = WScript.StdIn.ReadLine
	
	WScript.Echo vbLf & "Corresponding group in Active Directory : "
	group = WScript.StdIn.ReadLine
	
	WScript.Echo vbLf
	
	'Validate data
	If validateIndex(pronote, activeDirectory) Then
	
		groupDirectory = getGroupPath(group)
	
		Set OUgroup = getActiveOUDDirectoryFromRaw(groupDirectory, groupDirectory)
		
		Set OUdirectory = getActiveOUDirectory(activeDirectory)
		
		'Check that OU directory is valid (so function does not return null)
		If Not OUdirectory Is Nothing And Not OUgroup Is Nothing Then
			
			'Adding directory to the XML file
			Set xmlDoc = _
			CreateObject("Microsoft.XMLDOM")
			
			xmlDoc.Async = "False"
			'xmlDoc.indent = True
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
			
			Set objFieldValue = _
			  xmlDoc.createElement("Group")
			objFieldValue.Text = group
			objRecord.appendChild objFieldValue
			  
			'Save the file
			saveXML(xmlDoc)
			
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
		Call saveXML(xmlDoc)
		
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
	Set group = CreateObject("System.Collections.ArrayList")
	
	'Unique list in order to allow search functions to foreach only once
	Set uniqueActiveDirectory = CreateObject("System.Collections.ArrayList")
	
	'Foreach XML file
	For Each Index In xmlDoc.documentElement.SelectNodes("//Index")
	
		a = LCase(Index.getElementsByTagName("ActiveDirectory")(0).text)
		'Add to the list the AD index
		activeDirectory.Add a
		pronote.Add Index.getElementsByTagName("Pronote")(0).text
		group.Add LCase(Index.getElementsByTagName("Group")(0).text)
		
		If Not uniqueActiveDirectory.Contains(a) Then
			uniqueActiveDirectory.Add a
		End If
	Next
	
	'Return dictionnary
	Set d= CreateObject("Scripting.Dictionary")
	d.Add "activeDirectory", activeDirectory
	d.Add "pronote", pronote
	d.Add "group", group
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
	
	offset = maxSize - Len(text)
	
	If offset > 0 Then
		For i = 0 To offset - 1 Step 1
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

Sub saveXML(xmlDoc)
		set rdr = CreateObject("MSXML2.SAXXMLReader")
		set wrt = CreateObject("MSXML2.MXXMLWriter")
		Set oStream = CreateObject("ADODB.STREAM")
		oStream.Open
		oStream.Charset = "ISO-8859-1"
		 
		wrt.indent = True
		wrt.encoding = "ISO-8859-1"
		wrt.output = oStream
		Set rdr.contentHandler = wrt
		Set rdr.errorHandler = wrt
		rdr.Parse xmlDoc
		wrt.flush
		 
		oStream.SaveToFile xmlpath, 2
		 
		Set rdr = Nothing
		Set wrt = Nothing
End Sub

'From a raw path get a friednly path
Function getFriendlyPathFromOU(rawPath)
	parts = Split(rawPath, ",")

	friendlyPath = ""
	
	For Each p In parts 
	
		'Detect both uppercase and lowercase
		posUP = InStr(p, "OU=")
		posLO = InStr(p, "ou=")
		
		If posUP <> 0 Then
			friendlyPath =  Mid(p, posUP+3, Len(p)) & "/" & friendlyPath
		Else If posLO <> 0 Then
			friendlyPath = Mid(p, posLO+3, Len(p)) & "/" & friendlyPath
			End If
		End If
	Next
	
	'In case of wrong path, avoid program to crash
	If Len(friendlyPath) > 2 Then
		getFriendlyPathFromOU =  Mid(friendlyPath, 1, Len(friendlyPath)-1) 'Remove last "/"
	Else
		getFriendlyPathFromOU = ""
	End If
End Function

'Display table based on an array list of array list
Sub displayTable(rows)

	If rows.Count <= 1 Then
		Exit Sub
	End If

	'Compute length to know max length of every columns
	Set maxLength = CreateObject("System.Collections.ArrayList")
	
	'Init arrayList. Suppose that every lines have the same length
	For i = 0 To rows.Item(0).Count - 1 Step 1
		maxLength.Add 0 
	Next
	
	Dim pos
	For Each column In rows
		For i = 0 To column.Count - 1 Step 1
			maxLength.Item(i) = max(maxLength.Item(i), Len(column.Item(i)))
		Next
	Next
	
	'Get table border
	size = 0 
	For Each m In maxLength
		size = size + m
	Next
	border = generateTextBorder(size + 5 * rows.Item(0).Count)

	WScript.Echo border
	Dim display
	For Each column In rows
		display = "| "
		For i = 0 To column.Count - 1 Step 1
			display = display & column.Item(i) & getTableSpace(column.Item(i), maxLength.Item(i) + 2) & " | "
		Next
		WScript.Echo display
		WScript.Echo border
	Next

End Sub

'Clear arrayList and init header
Sub generateHeaders()

	textLogMoved.clear
	're add headers
	Set textLogMovedHeaders = CreateObject("System.Collections.ArrayList") 
	textLogMovedHeaders.Add "Student"
	textLogMovedHeaders.Add "From"
	textLogMovedHeaders.Add "To"
	textLogMoved.Add textLogMovedHeaders
	
	textLogCreated.clear
	're add headers
	Set textLogCreatedHeaders = CreateObject("System.Collections.ArrayList") 
	textLogCreatedHeaders.Add "Student"
	textLogCreatedHeaders.Add "Path"
	textLogCreatedHeaders.Add "Login"
	textLogCreated.Add textLogCreatedHeaders
	
	textLogUpdated.clear
	're add headers
	Set textLogUpdatedHeaders = CreateObject("System.Collections.ArrayList") 
	textLogUpdatedHeaders.Add "Student"
	textLogUpdatedHeaders.Add "FirstName"
	textLogUpdatedHeaders.Add "LastName"
	textLogUpdatedHeaders.Add "Email"
	textLogUpdatedHeaders.Add "Class"
	textLogUpdatedHeaders.Add "Group"
	textLogUpdated.Add textLogUpdatedHeaders
End Sub


'-------------------------------------------------------------------------------------------------------------
' END UTILITIES FUNCTIONS