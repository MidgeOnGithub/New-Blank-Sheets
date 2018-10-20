Attribute VB_Name = "NewBlankShts"
'New Blank Sheets [NBS] Version 1.0

'AUTHOR: Cody M Mason https://github.com/MidgeOnGithub

'LICENSE:

'The MIT License (MIT)

'Copyright (c) 2018 Cody M Mason

'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit

Sub NewBlankShts()
    Application.ScreenUpdating = False 'Speeds up code slightly
    
    Dim TWB As Workbook: Set TWB = ThisWorkbook 'Shorthand
    Dim WSCount As Integer: WSCount = TWB.Worksheets.Count 'Shorthand
    Dim i As Integer 'Loop incrementer
    Dim iBlanks As Integer 'Stores blank sheet count
    Dim arrBlanks() As Variant: ReDim arrBlanks(0) 'Array of blank sheet indices
    Dim sBlanks() As String: ReDim sBlanks(0) 'String array of blank worksheet Names (not CodeNames)
    Dim sCopies() As String: ReDim sCopies(0) 'String array of blank worksheets after being copied
    
    'Count hidden sheets with names starting with "Blank" and store info about them
    For i = 1 To WSCount
        With TWB.Worksheets(i)
            If InStr(.CodeName, "Blank") = 1 And Not .Visible = True Then
                iBlanks = iBlanks + 1
            
                If iBlanks > 1 Then
                    ReDim Preserve sBlanks(UBound(sBlanks) + 1)
                    ReDim Preserve sCopies(UBound(sCopies) + 1)
                    ReDim Preserve arrBlanks(UBound(arrBlanks) + 1)
                End If
            
                sBlanks(i - 1) = .Name
                sCopies(i - 1) = sBlanks(i - 1) & " (2)" 'Provides the sheet's post-copy name
                arrBlanks(i - 1) = i
            End If
        End With
    Next i
    
    Sheets(sBlanks).Copy After:=Sheets(arrBlanks(UBound(arrBlanks)))
     
    Dim iObjShts As Integer: iObjShts = 0 'Counts "Excel Object" modules with "Sheet" in CodeName
    
    For i = 1 To WSCount
        If InStr(TWB.Worksheets(i).CodeName, "Sheet") = 1 Then
            iObjShts = iObjShts + 1 'Count modules in TWB VBAProject with CodeNames starting with "Sheet"
        End If
    Next i
    
    Dim Now As String: Now = Format(DateTime.Now, "mmmdd HH.mm.ss") 'To name new sheets without conflict
    Dim sb0 As String 'Find Blank Sheet's object module using a leading 0 if applicable
    Dim s0 As String  'Names Sheet's object module with a leading 0 if applicable
    
    For i = LBound(sCopies) To UBound(sCopies)
        sb0 = "": If (iBlanks + i + 1) < 10 Then sb0 = "0"
        s0 = "": If (iObjShts + i + 1) < 10 Then s0 = "0"
        
        sCopies(i) = Right(sCopies(i), Len(sCopies(i)) - 6) 'Removes "Blank" and "(2)",
        sCopies(i) = Left(sCopies(i), Len(sCopies(i)) - 4)
        sCopies(i) = Now & " " & sCopies(i) 'Places Now in front of name to prevent name conflicts
        
        With TWB.Worksheets(arrBlanks(UBound(arrBlanks)) + i + 1) 'Add 1 because sCopies is indexed from 0, sheets from 1
            .Visible = True
            .Name = Left(sCopies(i), 31) 'Truncates name to max size if needed
        End With
        
        TWB.VBProject.VBComponents("Blank" & sb0 & (iBlanks + i + 1)).Name = "Sheet" & s0 & (iObjShts + i + 1)
    Next i
    
    Application.ScreenUpdating = True 'Return application to normal state
    
End Sub
