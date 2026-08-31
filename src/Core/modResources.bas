Attribute VB_Name = "modResources"
'@Folder("TPD_Addin.Core")

Public Sub CreateResourcesSheetIfMissing()
    Dim ws As Worksheet
    Dim logoPath As String
    Dim shp As Shape
    
    ' Path to PNG during development
    ' (Once embedded, this will no longer be needed)
    logoPath = "C:\dev\tpdLogo.png"   ' <-- update this once
    
    ' Check if _Resources already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Resources")
    On Error GoTo 0
    
    If Not ws Is Nothing Then Exit Sub   ' Already exists
    
    ' Validate PNG exists before creating the sheet
    If Dir(logoPath) = vbNullString Then
        MsgBox "Default logo file not found:" & vbCrLf & logoPath, vbExclamation, "Logo Missing"
        Exit Sub
    End If
    
    ' Create the sheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.name = "_Resources"
    
    ' Insert the PNG
    Set shp = ws.Shapes.AddPicture( _
        Filename:=logoPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=10, Top:=10, Width:=-1, Height:=-1)
    
    shp.name = "DefaultLogo"
    shp.LockAspectRatio = msoTrue
    
    ' Hide the sheet
    ws.Visible = xlSheetVeryHidden
End Sub


