Attribute VB_Name = "Module3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+q

    Set regEx = CreateObject("VBScript.RegExp")
    Dim strPattern As String: strPattern = "(([0-9]+)[/]([0-9]+))"
    Dim strReplace As String: strReplace = "$1"
    Dim strInput As String
    Dim Myrange As Range
    Set Myrange = ActiveSheet.UsedRange
    
    For Each cell In Myrange
        If strPattern <> "" Then
            strInput = cell.Value
            
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With
            
            If regEx.Test(strInput) Then
                    Set matches = regEx.Execute(strInput)
                    cell.Value = (regEx.Replace(strInput, Hi(matches(0).Value)))
            End If
        End If
    Next
End Sub

Function Hi(TestHi As String) As String

' Input: this function receives a string as a parameter
' Output: returns a string
Dim result As String
For l = 1 To Len(TestHi)
Dim den As Integer
Select Case True
    Case Mid(TestHi, l, 1) = "1"
        result = result & Chr(185)
    Case Mid(TestHi, l, 1) = "2"
        result = result & Chr(178)
    Case Mid(TestHi, l, 1) = "3"
        result = result & Chr("179")
    Case Mid(TestHi, l, 1) = "4"
        result = result & ChrW(&H2074)
    Case Mid(TestHi, l, 1) = "5"
        result = result & ChrW(&H2075)
    Case Mid(TestHi, l, 1) = "6"
        result = result & ChrW(&H2076)
    Case Mid(TestHi, l, 1) = "7"
        result = result & ChrW(&H2077)
    Case Mid(TestHi, l, 1) = "8"
        result = result & ChrW(&H2078)
    Case Mid(TestHi, l, 1) = "9"
        result = result & ChrW(&H2079)
    Case Mid(TestHi, l, 1) = "0"
        result = result & ChrW(&H2070)
    Case Mid(TestHi, l, 1) = "/"
        result = result & ChrW(&H2044)
        den = l
        Exit For
    End Select
Next


For l = den + 1 To Len(TestHi)
result = result & ChrW("&H208" & Mid(TestHi, l, 1))
Next

Hi = result


End Function



