Private Sub getMostRecentFile(path As String)
    Dim fileSystem As Object
    Dim myFolder As Folder
    Dim myFile As File
    Dim lastFileDate As Date
    Dim fileName, fullPath As String
    Dim targetWorkbook As Workbook
    
    'Creating the file system and assigning it a folder to work with
    Set fileSystem = CreateObject("Scripting.FilesystemObject")
    Set myFolder = fileSystem.GetFolder(path)
    
    'Looping through the folder, if the file has the latest date modified then capture the file
    lastFileDate = DateSerial(1900, 1, 1)
    For Each myFile In myFolder.Files
        If myFile.DateLastModified > lastFileDate Then
            fileName = myFile.Name
            lastFileDate = myFile.DateLastModified
        End If
    Next myFile
    
    'opening file
    fullPath = path & "\" & fileName
    Set targetWorkbook = Workbooks.Open(fullPath)
    
    'dereferencing objects
    Set fileSystem = Nothing
    Set myFolder = Nothing
End Sub
