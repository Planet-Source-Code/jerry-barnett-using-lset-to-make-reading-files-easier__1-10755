Attribute VB_Name = "UDTTestProj"
Option Explicit


Public Type udtInput
    MyDate     As String * 8
    Name       As String * 21
    Amount     As String * 7
    Code       As String * 1
    Account    As String * 6
End Type

Public Type udtLine
    strBuff    As String * 43
End Type


Private Sub Main()

Dim iMyFile As Long

Dim sInput As udtInput
Dim sLine As udtLine

Debug.Print LenB(sInput)

' Get File Number
iMyFile = FreeFile

' Open your file for reading
Open App.Path & "\SOMEFILE.TXT" For Input Access Read As #iMyFile

' Read the line into the udtLine UDT strBuff element
' This needs to be done, because you can't place a string
' directly into the UDT or you will get a type missmatch errorLine
Input #iMyFile, sLine.strBuff

' The buffer (strBuff) represents the entire line
' Now copy the udtLine UDT (sLine) into the udtInput
' UDT (sInput) *** ONE LINE OF CODE! ***
LSet sInput = sLine

Debug.Print sInput.MyDate

' Wolla! You can now access each element of the sInput UDT!
Debug.Print sInput.Name

' Will print: Jerry M Barnett
' (with five trailing spaces)
' Convert numeric String Amount value to a Long value with 2 decimal places
Debug.Print CDbl(Val(sInput.Amount) / 100)

' Would print: 23.56
'Debug.Print MyDateFunction(sInput.MyDate)

' Will print: 08072000 in any format you wish
' Note - The MyDateFunction is a function you define
' to parse the date string to the proper format you want.

Close iMyFile

End Sub
