'=========================================================================
'Queries a characters folder and search for attributes in the attributes.txt
'You must specify the attribute that you needed to search in getPropertyValue function
'=========================================================================

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim oFolder, folder, oSubFolder, sourceFolder, FSO, folderName, scriptDir

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Path: Set Path = FSO.GetFile(WScript.ScriptFullName)
scriptDir = Path.ParentFolder.ShortPath & "\"

Set oFolder = FSO.GetFolder(scriptDir)
Set oSubFolder = oFolder.SubFolders

If FSO.FileExists (scriptDir & "query.log") Then
	FSO.DeleteFile (scriptDir & "query.log")
End If

List "-------------------------"
List "-Querrying Characters...-"
List "-------------------------"


For Each folder in oSubFolder
	'List Characters
	List "Character Name: " & getFolderTitle(folder)
	List "Strength: " & getPropertyValue(folder, "str")
	List "Weapon 1: " & getPropertyValue(folder, "weapon1")
	List ""
Next

Wscript.Quit(0)


Sub List(strList)
	Dim ListFile
	
	If FSO.FileExists (scriptDir + "query.log") Then
		Set ListFile = FSO.OpenTextFile(scriptDir + "query.log", ForAppending, True)
	Else
		FSO.CreateTextFile scriptDir + "query.log", True
		Set ListFile = FSO.OpenTextFile(scriptDir + "query.log", ForAppending, True)
	End If
	ListFile.WriteLine strList
	ListFile.Close
End Sub

Function getPropertyValue (strLocation, strAttr)
	Dim strToSearch, strName, arrNames, intIndex
	strAttr = strAttr & "="
	strToSearch = strAttr
	Set objTextFile = FSO.OpenTextFile(strLocation & "\" & "Attributes.txt", ForReading)
	
	Do Until objTextFile.AtEndOfStream
		strLine = LCase(objTextFile.ReadLine())
		If InStr(strLine, strToSearch) <> 0 Then
			arrNames = Split(strLine, "=")
			intIndex = LBound(arrNames)
			
			If StrComp(strAttr, arrNames(intIndex) & "=") = 0 Then
				intIndex = UBound(arrNames)
				strName = arrNames(intIndex)
				Exit Do
			End If
		End If
	Loop
	objTextFile.Close
	getPropertyValue = LCase(strName)
End Function

Function getFolderTitle (strLocation)
	Dim arrNames, intIndex
	
	arrNames = Split(strLocation, "\")
	intIndex = UBound(arrNames)
	getFolderTitle = arrNames(intIndex)
End Function