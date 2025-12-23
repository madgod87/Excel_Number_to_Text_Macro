Sub ConvertToRupees()
    Dim srcRng As Range
    Dim destRng As Range
    Dim i As Long
    
    ' 1. Ask to select cell or cells with number
    On Error Resume Next
    Set srcRng = Application.InputBox("Select the cell(s) containing numbers to convert:", "Select Source Range", Type:=8)
    On Error GoTo 0
    
    If srcRng Is Nothing Then Exit Sub
    
    ' 2. Ask to select destination cell or cells
    On Error Resume Next
    Set destRng = Application.InputBox("Select the destination cell(s):", "Select Destination Range", Type:=8)
    On Error GoTo 0
    
    If destRng Is Nothing Then Exit Sub
    
    ' 3. Check if the number of cells are the same
    If srcRng.Cells.Count <> destRng.Cells.Count Then
        MsgBox "Error: The number of source cells (" & srcRng.Cells.Count & ") " & _
               "does not match the number of destination cells (" & destRng.Cells.Count & ").", vbCritical, "Range Mismatch"
        Exit Sub
    End If
    
    ' Process the conversion for each cell
    Application.ScreenUpdating = False
    For i = 1 To srcRng.Cells.Count
        If IsNumeric(srcRng.Cells(i).Value) And Not IsEmpty(srcRng.Cells(i)) Then
            destRng.Cells(i).Value = RupeeFormat(CStr(srcRng.Cells(i).Value))
        Else
            destRng.Cells(i).Value = ""
        End If
    Next i
    Application.ScreenUpdating = True
    
    MsgBox "Conversion completed successfully!", vbInformation
End Sub

' 4. Excel global function (UDF) facility
' Usage in sheet: =CONVERTTOTEXT(E3)
Public Function CONVERTTOTEXT(TargetCell As Variant) As String
    Dim inputVal As Variant
    
    ' Handle range input
    If TypeName(TargetCell) = "Range" Then
        inputVal = TargetCell.Value
    Else
        inputVal = TargetCell
    End If
    
    ' Validate input
    If Not IsNumeric(inputVal) Or IsEmpty(inputVal) Or inputVal = "" Then
        CONVERTTOTEXT = ""
        Exit Function
    End If
    
    ' Convert to word format
    CONVERTTOTEXT = RupeeFormat(CStr(inputVal))
End Function


Function RupeeFormat(SNum As String)
'Updateby Extendoffice
Dim xDPInt As Integer
Dim xArrPlace As Variant
Dim xRStr_Paisas As String
Dim xNumStr As String
Dim xF As Integer
Dim xTemp As String
Dim xStrTemp As String
Dim xRStr As String
Dim xLp As Integer
xArrPlace = Array("", "", " Thousand ", " Lacs ", " Crores ", " Trillion ", "", "", "", "")
On Error Resume Next
If SNum = "" Then
 RupeeFormat = ""
 Exit Function
End If
xNumStr = Trim(Str(SNum))
If xNumStr = "" Then
 RupeeFormat = ""
 Exit Function
End If
xRStr = ""
xLp = 0
If (Val(xNumStr) > 999999999.99) Then
   RupeeFormat = "Digit excced Maximum limit"
   Exit Function
End If
xDPInt = InStr(xNumStr, ".")
If xDPInt > 0 Then
   If (Len(xNumStr) - xDPInt) = 1 Then
      xRStr_Paisas = RupeeFormat_GetT(Left(Mid(xNumStr, xDPInt + 1) & "0", 2))
   ElseIf (Len(xNumStr) - xDPInt) > 1 Then
      xRStr_Paisas = RupeeFormat_GetT(Left(Mid(xNumStr, xDPInt + 1), 2))
   End If
   xNumStr = Trim(Left(xNumStr, xDPInt - 1))
End If
xF = 1
Do While xNumStr <> ""
   If (xF >= 2) Then
       xTemp = Right(xNumStr, 2)
   Else
       If (Len(xNumStr) = 2) Then
           xTemp = Right(xNumStr, 2)
       ElseIf (Len(xNumStr) = 1) Then
           xTemp = Right(xNumStr, 1)
       Else
           xTemp = Right(xNumStr, 3)
       End If
   End If
   xStrTemp = ""
   If Val(xTemp) > 99 Then
       xStrTemp = RupeeFormat_GetH(Right(xTemp, 3), xLp)
       If Right(Trim(xStrTemp), 3) <> "Lac" Then
       xLp = xLp + 1
       End If
   ElseIf Val(xTemp) <= 99 And Val(xTemp) > 9 Then
       xStrTemp = RupeeFormat_GetT(Right(xTemp, 2))
   ElseIf Val(xTemp) < 10 Then
       xStrTemp = RupeeFormat_GetD(Right(xTemp, 2))
   End If
   If xStrTemp <> "" Then
       xRStr = xStrTemp & xArrPlace(xF) & xRStr
   End If
   If xF = 2 Then
       If Len(xNumStr) = 1 Then
           xNumStr = ""
       Else
           xNumStr = Left(xNumStr, Len(xNumStr) - 2)
       End If
  ElseIf xF = 3 Then
       If Len(xNumStr) >= 3 Then
            xNumStr = Left(xNumStr, Len(xNumStr) - 2)
       Else
           xNumStr = ""
       End If
   ElseIf xF = 4 Then
     xNumStr = ""
Else
   If Len(xNumStr) <= 2 Then
   xNumStr = ""
Else
   xNumStr = Left(xNumStr, Len(xNumStr) - 3)
   End If
End If
   xF = xF + 1
Loop
If xRStr = "" Then
  xRStr = "No Rupees"
Else
  xRStr = " Rupees " & xRStr
End If
If xRStr_Paisas <> "" Then
  xRStr_Paisas = " and " & xRStr_Paisas & " Paisas"
End If
RupeeFormat = xRStr & xRStr_Paisas & " Only"
End Function
Function RupeeFormat_GetH(xStrH As String, xLp As Integer)
Dim xRStr As String
If Val(xStrH) < 1 Then
   RupeeFormat_GetH = ""
   Exit Function
Else
  xStrH = Right("000" & xStrH, 3)
  If Mid(xStrH, 1, 1) <> "0" Then
       If (xLp > 0) Then
        xRStr = RupeeFormat_GetD(Mid(xStrH, 1, 1)) & " Lac "
       Else
        xRStr = RupeeFormat_GetD(Mid(xStrH, 1, 1)) & " Hundred "
       End If
   End If
   If Mid(xStrH, 2, 1) <> "0" Then
       xRStr = xRStr & RupeeFormat_GetT(Mid(xStrH, 2))
   Else
       xRStr = xRStr & RupeeFormat_GetD(Mid(xStrH, 3))
   End If
End If
   RupeeFormat_GetH = xRStr
End Function
Function RupeeFormat_GetT(xTStr As String)
   Dim xTArr1 As Variant
   Dim xTArr2 As Variant
   Dim xRStr As String
   xTArr1 = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
   xTArr2 = Array("", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
   result = ""
   If Val(Left(xTStr, 1)) = 1 Then
       xRStr = xTArr1(Val(Mid(xTStr, 2, 1)))
   Else
       If Val(Left(xTStr, 1)) > 0 Then
           xRStr = xTArr2(Val(Left(xTStr, 1)) - 1)
       End If
       xRStr = xRStr & RupeeFormat_GetD(Right(xTStr, 1))
   End If
     RupeeFormat_GetT = xRStr
End Function
Function RupeeFormat_GetD(xDStr As String)
Dim xArr_1() As Variant
   xArr_1 = Array(" One", " Two", " Three", " Four", " Five", " Six", " Seven", " Eight", " Nine", "")
   If Val(xDStr) > 0 Then
       RupeeFormat_GetD = xArr_1(Val(xDStr) - 1)
   Else
       RupeeFormat_GetD = ""
   End If
End Function


