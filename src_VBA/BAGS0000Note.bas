Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'This version (Version 2007) was modified in August 2007 to include adjustment to roughness.

'Graphics for sediment transport, transport stage and water depth rating curves are
'added during the modification.

'August 2006: Parker (1990) more transport by excluding 2 mm and finer. -- need some attention!

'November 2008:
'   a). Check freezing problem when enter width or cross-section - Fixed.
'   b). Some students reported that PK does not run. - Cannot find it, seems unlikely because
'       PK is the simplest equation of all, so must be an input error.
'   b). If Qmax too large, running a sediment rating curve leads to Excel not responding. --
'       A label is added to forewarn users to chose a reasonable maximum discharge.

Option Explicit

Function ULCase(MyString As String)
    Dim cc As String
    Dim i As Long
    
    ULCase = ""
    For i = 1 To Len(MyString)
        cc = Mid(MyString, i, 1)
        If i = 1 Then
            cc = UCase(cc)
        ElseIf Mid(MyString, i - 1, 1) = " " Then
            cc = UCase(cc)
        End If
        ULCase = ULCase & cc
    Next
End Function


Sub ViewAll()
    Dim i As Long
    ActiveWindow.DisplayWorkbookTabs = True
    For i = 1 To Sheets.Count
        Sheets(i).Visible = True
    Next
End Sub


