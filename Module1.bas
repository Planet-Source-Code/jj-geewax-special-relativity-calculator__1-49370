Attribute VB_Name = "Module1"
Option Explicit

Public Function CalcXPrime(X As Double, v As Double, T As Double) As Double
On Error GoTo ErrorX:
    CalcXPrime = Gamma(v) * (X - (v * T))
Exit Function
ErrorX:
CalcXPrime = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcX(XPrime As Double, v As Double, TPrime As Double) As Double
On Error GoTo ErrorX:
    CalcX = Gamma(v) * (XPrime + (v * TPrime))
Exit Function
ErrorX:
CalcX = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function Gamma(v As Double) As Double
On Error GoTo ErrorX:
Dim Dummy As Double
    Dummy = (1 - (v ^ 2))
    If Not Dummy = 0 Then
        Gamma = 1 / Sqr(Dummy)
    Else
        Gamma = 0
    End If
Exit Function
ErrorX:
Gamma = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcL(LPrime As Double, v As Double) As Double
On Error GoTo ErrorX:
Dim Dummy As Double
    Dummy = Sqr(1 - (v ^ 2))
    'we dont need to cancel the c's out because they arent there.
    CalcL = Dummy * LPrime
Exit Function
ErrorX:
CalcL = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcLPrime(L As Double, v As Double) As Double
On Error GoTo ErrorX:
    CalcLPrime = Gamma(v) * L
Exit Function
ErrorX:
CalcLPrime = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcDeltaT(DeltaTPrime As Double, v As Double) As Double
On Error GoTo ErrorX:
    CalcDeltaT = Gamma(v) * DeltaTPrime
Exit Function
ErrorX:
CalcDeltaT = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcDeltaTPrime(DeltaT As Double, v As Double) As Double
On Error GoTo ErrorX:
Dim Dummy As Double
    Dummy = Sqr(1 - (v ^ 2))
    CalcDeltaTPrime = Dummy * DeltaT
Exit Function
ErrorX:
CalcDeltaTPrime = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcTPrime(T As Double, v As Double, X As Double) As Double
On Error GoTo ErrorX:
    CalcTPrime = Gamma(v) * (T - (v * X))
Exit Function
ErrorX:
CalcTPrime = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcT(TPrime As Double, v As Double, XPrime As Double) As Double
On Error GoTo ErrorX:
    CalcT = Gamma(v) * (TPrime + (v * XPrime))
Exit Function
ErrorX:
CalcT = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcUXPrime(ux As Double, v As Double) As Double
On Error GoTo ErrorX:
Dim Dummy2 As Double
    Dummy2 = (1 - (v * ux))
    If Not Dummy2 = 0 Then
        CalcUXPrime = (ux - v) / Dummy2
    Else
        CalcUXPrime = 0#
    End If
Exit Function
ErrorX:
CalcUXPrime = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcUX(uxprime As Double, v As Double) As Double
On Error GoTo ErrorX:
Dim Dummy3 As Double
    Dummy3 = (1 + (v * uxprime))
    If Not Dummy3 = 0 Then
        CalcUX = (uxprime + v) / Dummy3
    Else
        CalcUX = 0#
    End If
Exit Function
ErrorX:
CalcUX = 0
MsgBox "An error occurred! Check you values."
End Function

Public Function CalcVRel(ux As Double, uxprime As Double) As Double
On Error GoTo ErrorX:
Dim Dummy4 As Double
Dummy4 = 1 - (uxprime * ux)
If Not Dummy4 = 0 Then
    CalcVRel = (ux - uxprime) / Dummy4
Else
    CalcVRel = 0
End If
CalcVRel = Round(CalcVRel, 3)
Exit Function
ErrorX:
CalcVRel = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcVLenCont(L As Double, LPrime As Double) As Double
On Error GoTo ErrorX:
    Dim Dummy5 As Double
    If Not LPrime = 0 Then
        Dummy5 = ((L) / (LPrime)) ^ 2
        CalcVLenCont = (1 - Dummy5)
        CalcVLenCont = Sqr(CalcVLenCont)
    Else
        CalcVLenCont = 0
    End If
Exit Function
ErrorX:
CalcVLenCont = 0
MsgBox "An error occurred! Check you values."
End Function
Public Function CalcVTimeDil(T As Double, TPrime As Double) As Double
On Error GoTo ErrorX:
    Dim Dummy6 As Double
    Dummy6 = T * T
    If Not Dummy6 = 0 Then
        CalcVTimeDil = Sqr(1 - ((TPrime ^ 2) / Dummy6))
    Else
        CalcVTimeDil = 0
    End If
Exit Function
ErrorX:
CalcVTimeDil = 0
MsgBox "An error occurred! Check you values."
End Function
Sub SaveText(txtSave As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Sub SaveList(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub Loadlist(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
