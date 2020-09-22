VERSION 5.00
Begin VB.UserControl PBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ForeColor       =   &H00000000&
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ToolboxBitmap   =   "PBar.ctx":0000
End
Attribute VB_Name = "PBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''
''ProgressBar Replacement By Daniel Taylor'''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Copyright 2001''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''

Public Enum Style_Type
    Out
    Etch
    DottedLine
    Plain
    None
End Enum

Public Enum Align_Type
    Left_Align
    Center_Align
    Right_Align
End Enum

Public Enum Bar_End_Type
    Normal
    Tip
    Arrow
End Enum

'My Variables
Dim BorderWidth As Integer
Dim BarEndLength As Integer

'Standard-Eigenschaftswerte:
Const m_def_BarEnd = 0
Const m_def_Text = ""
Const m_def_TextAlign = 1
Const m_def_BarLength = -1
Const m_def_ReverseColor = 1
Const m_def_ShowPercent = 1
Const m_def_ForeColor = &HF86832
Const m_def_Value = 100
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_BorderStyle = 1

'Eigenschaftsvariablen:
Dim m_BarEnd As Bar_End_Type
Dim m_Text As String
Dim m_TextAlign As Align_Type
Dim m_BarLength As Integer
Dim m_ReverseColor As Boolean
Dim m_ShowPercent As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_Value As Integer
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_BorderStyle As Style_Type

'API
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Events
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MaxReached()
Event ValueChanged(Value As Integer)
Event RefreshFinish()

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,1
Public Property Get ShowPercent() As Boolean
    ShowPercent = m_ShowPercent
End Property

Public Property Let ShowPercent(ByVal New_ShowPercent As Boolean)
    m_ShowPercent = New_ShowPercent
    PropertyChanged "ShowPercent"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,&H80000012&
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TextColor = UserControl.ForeColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    UserControl.ForeColor() = New_TextColor
    PropertyChanged "TextColor"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,100
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    PropertyChanged "Value"
    Refresh
    RaiseEvent ValueChanged(Value)
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
    PropertyChanged "Min"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    PropertyChanged "Max"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,0
Public Property Get BorderStyle() As Style_Type
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Style_Type)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Refresh
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_ShowPercent = m_def_ShowPercent
    m_ForeColor = m_def_ForeColor
    m_Value = m_def_Value
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_BorderStyle = m_def_BorderStyle
    Set UserControl.Font = Ambient.Font
    Refresh
    m_ReverseColor = m_def_ReverseColor
    m_BarLength = m_def_BarLength
    m_Text = m_def_Text
    m_TextAlign = m_def_TextAlign
    m_BarEnd = m_def_BarEnd
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ShowPercent = PropBag.ReadProperty("ShowPercent", m_def_ShowPercent)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("TextColor", &H80000012)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ReverseColor = PropBag.ReadProperty("ReverseColor", m_def_ReverseColor)
    m_BarLength = PropBag.ReadProperty("BarLength", m_def_BarLength)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_BarEnd = PropBag.ReadProperty("BarEnd", m_def_BarEnd)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Show()
    Refresh
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ShowPercent", m_ShowPercent, m_def_ShowPercent)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("TextColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ReverseColor", m_ReverseColor, m_def_ReverseColor)
    Call PropBag.WriteProperty("BarLength", m_BarLength, m_def_BarLength)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("BarEnd", m_BarEnd, m_def_BarEnd)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Function Refresh()
    On Error Resume Next
    Dim Percent As Integer
    Percent = (Value / (Max - Min)) * 100
    If Percent >= 100 Then
        Percent = 100
    ElseIf Percent < 0 Then
        Percent = 0
    End If
    UserControl.Cls
    DrawBorder
    If Value > 0 Then
        GetBarEndWidth
        If BarLength = -1 Then
            UserControl.Line ((0 + BorderWidth) - BarEndLength, 0 + BorderWidth)-(((UserControl.ScaleWidth - (BorderWidth + 1)) * (Percent / 100)) - BarEndLength, (UserControl.ScaleHeight - (BorderWidth + 1))), ForeColor, BF
        Else
            UserControl.Line ((((UserControl.ScaleWidth + BorderWidth) * (Percent / 100)) - BarLength) - BarEndLength, 0 + BorderWidth)-(((UserControl.ScaleWidth - (BorderWidth + 1)) * (Percent / 100) - BarEndLength), (UserControl.ScaleHeight - (BorderWidth + 1))), ForeColor, BF
        End If
        DrawBarEnd (UserControl.ScaleWidth - (BorderWidth + 1)) * (Percent / 100)
    End If
    DrawText Percent, Text
    RaiseEvent RefreshFinish
    If Percent = 100 Then
        RaiseEvent MaxReached
    End If
End Function

Private Sub DrawBorder()
    Dim Color1 As OLE_COLOR, Color2 As OLE_COLOR
    Color1 = -1
    Color2 = -1
    CheckForColors Color1, Color2
    If ReverseColor = True Then
        Dim CHOLD As OLE_COLOR
        CHOLD = Color1
        Color1 = Color2
        Color2 = CHOLD
        CHOLD = Empty
    End If
    If BorderStyle = Out Then
        BorderWidth = 1
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), Color1
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 1), Color1
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Color2
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Color2
    ElseIf BorderStyle = Etch Then
        BorderWidth = 2
        'outside
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), Color2
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 1), Color2
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Color1
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Color1
        'inside
        UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2, 1), Color1
        UserControl.Line (1, 1)-(1, UserControl.ScaleHeight - 2), Color1
        UserControl.Line (1, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), Color2
        UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), Color2
    ElseIf BorderStyle = DottedLine Then
        Dim X As Integer, Y As Integer
        BorderWidth = 1
        For X = 0 To UserControl.ScaleWidth Step 2
            SetPixelV UserControl.hdc, X, 0, Color2
            SetPixelV UserControl.hdc, X, UserControl.ScaleHeight - 1, Color2
        Next X
        For Y = 0 To UserControl.ScaleHeight Step 2
            SetPixelV UserControl.hdc, 0, Y, Color2
            SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, Y, Color2
        Next Y
    ElseIf BorderStyle = Plain Then
        BorderWidth = 1
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Color2, B
    ElseIf BorderStyle = None Then
        borderwith = 0
    End If
End Sub

Private Sub DrawText(Percent As Integer, Optional TXT As String = "")
    If ShowPercent = True Or Text <> "" Then
        Dim X As Integer, Y As Integer, P As String
        If TXT = "" Then
            P = Percent & "%"
        Else
            If ShowPercent = True Then
                P = TXT & Percent & "%"
            Else
                P = TXT
            End If
        End If
        If TXT = "" Or TextAlign = Center_Align Then
            X = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(P) / 2)
            Y = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(P) / 2)
        ElseIf TextAlign = Left_Align Then
            X = 0 + BorderWidth + 1
            Y = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(P) / 2)
        Else
            X = (UserControl.ScaleWidth) - (UserControl.TextWidth(P) + BorderWidth + 1)
            Y = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(P) / 2)
        End If
        UserControl.CurrentX = X
        UserControl.CurrentY = Y
        UserControl.Print P
    End If
End Sub

Private Function GetRGB(Color As OLE_COLOR, Red, Green, Blue)
    TranslateColor Color, 0, Color
    Red = Color And &HFF
    Green = (Color And &HFF00&) / 255
    Blue = (Color And &HFF0000) / 65536
End Function

Private Function CheckForColors(Optional Color1 As OLE_COLOR = 0, Optional Color2 As OLE_COLOR = 0)
    If Color1 = -1 Or Color2 = -1 Then
        Dim R As Integer, G As Integer, B As Integer
        GetRGB BackColor, R, G, B
        If Color1 = -1 Then
            If R > 199 Then
                R = 255
            Else
                R = R + 50
            End If
            If G > 199 Then
                G = 255
            Else
                G = G + 50
            End If
            If B > 199 Then
                B = 255
            Else
                B = B + 50
            End If
            Color1 = RGB(R, G, B)
        End If
        If Color2 = -1 Then
            If R < 141 Then
                R = 0
            Else
                R = R - 140
            End If
            If G < 141 Then
                G = 0
            Else
                G = G - 140
            End If
            If B < 141 Then
                B = 0
            Else
                B = B - 140
            End If
            Color2 = RGB(R, G, B)
        End If
    End If
End Function
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get ReverseColor() As Boolean
    ReverseColor = m_ReverseColor
End Property

Public Property Let ReverseColor(ByVal New_ReverseColor As Boolean)
    m_ReverseColor = New_ReverseColor
    PropertyChanged "ReverseColor"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,-1
Public Property Get BarLength() As Integer
    BarLength = m_BarLength
End Property

Public Property Let BarLength(ByVal New_BarLength As Integer)
    m_BarLength = New_BarLength
    PropertyChanged "BarLength"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,
Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,0
Public Property Get TextAlign() As Align_Type
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As Align_Type)
    m_TextAlign = New_TextAlign
    PropertyChanged "TextAlign"
    Refresh
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,0
Public Property Get BarEnd() As Bar_End_Type
Attribute BarEnd.VB_Description = "Special Way the end of the loading bar looks"
    BarEnd = m_BarEnd
End Property

Public Property Let BarEnd(ByVal New_BarEnd As Bar_End_Type)
    m_BarEnd = New_BarEnd
    PropertyChanged "BarEnd"
    Refresh
End Property

Private Function GetBarEndWidth()
    If BarEnd = Normal Then
        BarEndLength = 0
    ElseIf BarEnd = Tip Then
        BarEndLength = UserControl.ScaleHeight
    ElseIf BarEnd = Arrow Then
        BarEndLength = UserControl.ScaleHeight
    Else
        BarEndLength = 10
    End If
End Function

Private Sub DrawBarEnd(StartFrom As Integer)
    StartFrom = StartFrom - BarEndLength
    If BarEnd = Tip Then
        UserControl.Line (StartFrom, BorderWidth)-(StartFrom + BarEndLength, UserControl.ScaleHeight / 2), ForeColor
        UserControl.Line (StartFrom, UserControl.ScaleHeight - (BorderWidth + 1))-(StartFrom + BarEndLength, UserControl.ScaleHeight / 2), ForeColor
    ElseIf BarEnd = Arrow Then
        UserControl.Line (StartFrom, (UserControl.ScaleHeight * 0.25) + BorderWidth)-(StartFrom + 4, (UserControl.ScaleHeight * 0.25) + BorderWidth), ForeColor
        UserControl.Line (StartFrom, (UserControl.ScaleHeight * 0.75) - BorderWidth)-(StartFrom + 4, (UserControl.ScaleHeight * 0.75) - BorderWidth), ForeColor
        StartFrom = StartFrom + 4
        BarEndLength = BarEndLength - 4
        UserControl.Line (StartFrom, (UserControl.ScaleHeight * 0.25) + BorderWidth)-(StartFrom, BorderWidth), ForeColor
        UserControl.Line (StartFrom, (UserControl.ScaleHeight * 0.75) - BorderWidth)-(StartFrom, UserControl.ScaleHeight - BorderWidth), ForeColor
        UserControl.Line (StartFrom, BorderWidth)-(StartFrom + BarEndLength, UserControl.ScaleHeight / 2), ForeColor
        UserControl.Line (StartFrom, UserControl.ScaleHeight - (BorderWidth + 1))-(StartFrom + BarEndLength, UserControl.ScaleHeight / 2), ForeColor
    End If
End Sub
