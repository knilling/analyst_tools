' BSD 2-Clause License
' 
' Copyright (c) 2019, Chris Crawford
' All rights reserved.
' 
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
' 
' * Redistributions of source code must retain the above copyright notice, this
'   list of conditions and the following disclaimer.
' 
' * Redistributions in binary form must reproduce the above copyright notice,
'   this list of conditions and the following disclaimer in the documentation
'   and/or other materials provided with the distribution.
' 
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Attribute VB_Name = "Analyst_Tools"
' Requires cStack

Private Sub InsertColumnToRight()
Attribute InsertColumnToRight.VB_ProcData.VB_Invoke_Func = " \n14"
    ' https://www.mrexcel.com/forum/excel-questions/77371-vba-insert-row-colum.html#post374234
    ActiveCell.EntireColumn.Offset(0, 1).Insert
End Sub

Private Sub InsertColumnToLeft()
    ' https://www.mrexcel.com/forum/excel-questions/77371-vba-insert-row-colum.html#post374234
    ActiveCell.EntireColumn.Insert
End Sub

Private Function getValue(c As Range)
    getValue = c.Cells(1, 1).value
End Function

Private Function GetRow(c As Range)
    GetRow = c.Cells(1, 1).Row
End Function

Private Function GetColumn(c As Range)
    GetColumn = c.Cells(1, 1).Column
End Function

Private Function isIPv4(addr)
    ' Regexes in VBA: https://stackoverflow.com/a/22542835
    ' Tools -> References -> Microsoft VBScript Regular Expressions 5.5
    Dim regEx As New RegExp
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.MultiLine = False
    ' Regex from: https://stackoverflow.com/a/5284410
    regEx.Pattern = "\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}\b"
    If regEx.Test(addr) Then
        isIPv4 = True
    Else
        isIPv4 = False
    End If
End Function

Private Sub splitIP(c As Range)
    myValue = getValue(c)
    myRow = GetRow(c)
    myCol = GetColumn(c)
    myValues = Split(myValue, ".")
    Excel.Cells(myRow, myCol + 1).value = myValues(0)
    Excel.Cells(myRow, myCol + 2).value = myValues(1)
    Excel.Cells(myRow, myCol + 3).value = myValues(2)
    Excel.Cells(myRow, myCol + 4).value = myValues(3)
End Sub

Private Sub insertFourColumns()
    InsertColumnToRight
    InsertColumnToRight
    InsertColumnToRight
    InsertColumnToRight
End Sub

Sub Split_IP_Addresses()
    Dim c As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    insertFourColumns
    For Each c In selectedRange.Cells
        my_val = getValue(c)
        If isIPv4(my_val) Then
            Call splitIP(c)
        End If
    Next c
End Sub

Function largest(c As Collection)
    n = 0
    For Each i In c
        If n < i Then
            n = i
        End If
    Next i
    largest = n
End Function

Sub Split_Domain_Names()
    Dim c As Range
    Dim l As Collection
    Dim selectedRange As Range
    Set l = New Collection
    Set selectedRange = Application.Selection
    For Each c In selectedRange.Cells
        my_val = getValue(c)
        Dim parts() As String
        parts = Split(my_val, ".")
        n = UBound(parts) + 1
        Call l.Add(n)
    Next c
    
    zulu = largest(l)
    zulu = zulu + 1
    
    For i = 1 To zulu
        InsertColumnToRight
    Next i
    
    For Each c In selectedRange.Cells
        my_val = getValue(c)
        Dim parts2() As String
        parts2 = Split(my_val, ".")
        n = UBound(parts2) + 1
        
        Dim cs As New cStack
        Set cs = New cStack
        cs.init
        For Each part In parts2
            cs.Push (part)
        Next part
        
        myRow = GetRow(c)
        myCol = GetColumn(c)
        Excel.Cells(myRow, myCol + 1).value = n
        
        anotherPart = cs.Pop
        my_i = 2
        While Not Len(anotherPart) = 0
            Excel.Cells(myRow, myCol + my_i).value = anotherPart
            anotherPart = cs.Pop
            my_i = my_i + 1
        Wend
    Next c
    
End Sub
