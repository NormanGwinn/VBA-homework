Attribute VB_Name = "mUtility"
Option Explicit

Public Function OpenWorkbook(sDirectory As String, sFilename As String) As Workbook
    Dim wbk As Workbook
    Dim sFullPath As String
    
    On Error Resume Next
    Set wbk = Application.Workbooks.Item(sFilename)
    If wbk Is Nothing Then
        If Right(sDirectory, 1) = "/" Then
            sFullPath = sDirectory & sFilename
        Else
            sFullPath = sDirectory & "/" & sFilename
        End If
        Set wbk = Application.Workbooks.Open(Filename:=sFullPath, ReadOnly:=True)
    End If
    On Error GoTo 0
    Set OpenWorkbook = wbk
End Function
