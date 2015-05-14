<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
	<title>New Plymouth Reformed Church Sermon Archive</title>
</head>
<body>
<h1>New Plymouth Reformed Church Sermon Archive</h1>
<!--Do Not Modify The Following Code, Only Application("ICONS_PATH"), and Application("ICONS_VIRTUAL_PATH") If Needed-->
<%
' Constants
Const FileAttrHidden = 2

' Settings
Dim currentPath, currentPage, iconsPath
currentPath = Split(Request.ServerVariables("PATH_TRANSLATED"), "\", -1, 1) ' An array of the path to this file
currentPage = currentPath(UBound(currentPath)) ' The name of this file

' Set the path to where the file icons are kept if it is empty
' Note: You should set these in global.asa
' All icons in this folder must be in gif format
' For example, the icon for Word documents would be doc.gif
If Application("ICONS_VIRTUAL_PATH") = "" Then Application("ICONS_VIRTUAL_PATH") = "/icons" ' Manually set the virtual path here if you don't have a global variable
If Application("ICONS_PATH") = "" Then Application("ICONS_PATH") = "C:\Inetpub\wwwroot\icons" ' Manually set the path here if you don't have a global variable

' Start Filesystem access
Set fs = CreateObject("Scripting.FileSystemObject")

' Display an up button. Remove this line if this file is in the root of your website
If fs.FileExists(Application("ICONS_PATH") + "\parent.gif") And UBound(currentPath) > 2 Then
	Response.Write("<a href=""../""><img src=""" + Application("ICONS_VIRTUAL_PATH") + "/parent.gif"" align=""absmiddle"" border=""0"" /></a> <a href=""../"">Parent Directory</a><br>")
End If

' Set the name of the current folder
Set folder = fs.GetFolder(Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(CurrentPage)))

' Get all of the folders in the current folder
Set folders = folder.SubFolders

' Display each folder one by one
For Each foldername In folders
	' If the file found is not this page and is not hidden or system, then display a link to it
	If Not foldername.Attributes And FileAttrHidden Then
		' Display the icon if one exists
		If fs.FileExists(Application("ICONS_PATH") + "\folder.gif") Then
			Response.Write("<a href=""" + foldername.Name + """><img src=""" + Application("ICONS_VIRTUAL_PATH") + "/folder.gif"" align=""absmiddle"" border=""0"" /></a>")
		End If
		' Display the name of the folder
		Response.Write(" <a href=""" + foldername.Name + "/"">" + foldername.Name + "</a><br>")
	End If
Next

' Get all of the files in the current folder
Set files = folder.Files

' Display each file one by one
For Each filename In files
	' If the file found is not this page and is not hidden or system, then display a link to it
	If filename.Name <> currentPage And (Not filename.Attributes And FileAttrHidden) Then
		' Display the icon if one exists. This If statement speeds processing.
		If Application("ICONS_PATH") <> "" Then
			' Get the file's extension
			Dim fileExtension, filenameArray
			filenameArray = Split(filename.Name, ".", -1, 1)
			fileExtension = filenameArray(UBound(filenameArray))
			' If we have an image for this filetype show it, otherwise shw the default icon if it exists
			If fs.FileExists(Application("ICONS_PATH") + "\" + fileExtension + ".gif") Then
				Response.Write("<a href=""" + filename.Name + """><img src=""" + Application("ICONS_VIRTUAL_PATH") + "/" + fileExtension + ".gif"" align=""absmiddle"" border=""0"" /></a>")
			ElseIf fs.FileExists(Application("ICONS_PATH") + "\default.gif") Then
				Response.Write("<a href=""" + filename.Name + """><img src=""" + Application("ICONS_VIRTUAL_PATH") + "/default.gif"" align=""absmiddle"" border=""0"" /></a>")
			End If
		End If
		' Display the name of the file
		Response.Write(" <a href=""" + filename.Name + """>" + filename.Name + "</a><br>")
	End If
Next
 %>
<!--You Can Modify All Of The Code Below-->
<p align="right">&copy; 2015 New Plymouth Reformed Church</p>
</body>

</html>

