Attribute VB_Name = "ProgressBarCircle"
Option Explicit

Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim r, g, b As Integer
Dim LastNum As Integer
Dim i As Long
Dim X, Y As Long
Dim SelCol As String
Dim RainBowCol As Boolean

Public C, S As Single
Public MeterValue As Single
Public SL, ST, Size As Integer
Public Num As Integer

Public Function LoadMeter(CurForm As Form, BarColor As Long, PercentColor As Long, RainBow As Boolean)
'To Make RainBow Call Either a R, G, B Color In BarColor
On Error GoTo Error
'Basic Numbers
SL = CurForm.MeterShape.Left + CurForm.MeterShape.Width / 2
ST = CurForm.MeterShape.Top + CurForm.MeterShape.Height / 2
Size = CurForm.MeterShape.Width / 2
CurForm.MeterBox.Circle (SL, ST), CurForm.MeterShape.Width / 2 + 1
CurForm.MeterBox.Picture = CurForm.MeterBox.Image
'MeterPos
CurForm.MeterPos.Caption = "0%"
CurForm.MeterPos.Left = 0
CurForm.MeterPos.Width = CurForm.MeterBox.ScaleWidth
CurForm.MeterPos.Top = CurForm.MeterShape.Top + CurForm.MeterShape.Height / 2 - CurForm.MeterPos.Height / 2
CurForm.MeterPos.ForeColor = PercentColor
'Meter
If RainBow = False Then
    CurForm.MeterBox.ForeColor = BarColor
Else
    RainBowCol = True
    r = BarColor Mod &H100
    g = (BarColor \ &H100) Mod &H100
    b = (BarColor \ &H10000) Mod &H100
    If r > g And r > b Then
        r = 1
        g = 0
        b = 0
        GoTo Finish
        End If
    If g > r And g > b Then
        r = 0
        g = 1
        b = 0
        GoTo Finish
        End If
    If b > r And b > g Then
        r = 0
        g = 0
        b = 1
        GoTo Finish
        End If
    If r = g Then
        r = 1
        g = 1
        b = 0
        GoTo Finish
        End If
    If r = b Then
        r = 1
        g = 0
        b = 1
        GoTo Finish
        End If
    If g = b Then
        r = 0
        g = 1
        b = 1
        GoTo Finish
        End If
End If
Finish:
SetMeter 0, Form1
Exit Function
Error:
MsgBox Err.Description
End Function

Public Function SetMeter(SetNum As Single, CurForm As Form)
On Error GoTo Error
Dim MeterNum As Single
'Checks if SetNum for Certain Values
SetNum = Round(SetNum)
Select Case SetNum
Case Is <= -3
    SetNum = 0
Case -1
    'Go Up One
    SetNum = LastNum + 1
    If SetNum >= 100 Then
        SetNum = 100
        End If
Case -2
    'Go Down One
    SetNum = LastNum - 1
    If SetNum < 0 Then
        SetNum = 0
        End If
Case Is > 100
    Exit Function
End Select

'Sets the Circle Bar
CurForm.MeterBox.Cls
For i = 0 To Round(SetNum * 3.6, 0)
    C = i
    S = i
    C = Cos(C * (3.14159 / 180))
    S = Sin(S * (3.14159 / 180))
    
    If RainBowCol = True Then
        CurForm.MeterBox.ForeColor = RGB(i / 1.4 * r, i / 1.4 * g, i / 1.4 * b)
        End If

    '***ClockWise Progress
    CurForm.MeterBox.Line (SL, ST)-(S * Size + SL, -C * Size + ST)
    
    '***Counter-ClockWise Progress
    'CurForm.MeterBox.Line (SL, ST)-(-S * Size + SL, -C * Size + ST)
    
    '***Upside Down ClockWise Progress
    'CurForm.MeterBox.Line (SL, ST)-(S * Size + SL, C * Size + ST)
    
    '***Upside Down Counter-ClockWise Progress
    'CurForm.MeterBox.Line (SL, ST)-(-S * Size + SL, C * Size + ST)
    
    '***45 Degrees Progress ClockWise
    'CurForm.MeterBox.Line (SL, ST)-(C * Size + SL, S * Size + ST)
    
    '***45 Degrees Progress Counter-ClockWise
    'CurForm.MeterBox.Line (SL, ST)-(C * Size + SL, -S * Size + ST)
    
    '***135 Degress Progress ClockWise
    'CurForm.MeterBox.Line (SL, ST)-(-C * Size + SL, -S * Size + ST)
    
    '***135 Degress Progress Counter-CloseWise
    'CurForm.MeterBox.Line (SL, ST)-(-C * Size + SL, S * Size + ST)
Next i
CurForm.MeterBox.Refresh
CurForm.MeterPos.Caption = SetNum & "%"
LastNum = SetNum
Exit Function
Error:
MsgBox Err.Description
End Function

Public Function GetMeter(CurForm As Form)
MeterValue = LastNum
End Function
