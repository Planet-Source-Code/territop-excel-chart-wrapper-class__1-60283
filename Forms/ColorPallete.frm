VERSION 5.00
Begin VB.Form ColorPallete 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excel Color Pallete"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2295
   Icon            =   "ColorPallete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
End
Attribute VB_Name = "ColorPallete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
'       CExcelPlot - Wrapper Class for  Microsoft Excel Charts using OLE2 Embedding
'
'   Product Name:
'       CExcelPlot.cls
'
'   Compatability:
'       Windows: 2000, XP, (Perhaps others, but not tested!)
'       Excel Version(Year): 7(1995), 8(1997), 9(2000), 10(2002), 11(2003)
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Legal Copyright & Trademarks:
'       Copyright © 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       20Feb05 - Initial test harness for CExcelPlot Wrapper Class finished
'
'   Force Declarations
Option Explicit
Private i               As Integer          'Loop Counter
Private j               As Integer          'Loop Counter
Private k               As Integer          'Loop Counter
Private xlColors()      As Variant          'Excel's Internal Color Table
Private m_PrevIndex     As Integer          'Previous Index Value
Private m_ColorIndex    As Long             'Color Index value from Calling form

Property Get ColorIndex() As Long
    ColorIndex = m_ColorIndex
End Property

Property Let ColorIndex(Index As Long)
        '   Check to see if we are within range
    If (Color >= 1) And (Color < 57) Then
        m_ColorIndex = Index
    Else
        '   We are not so tell the user
        Err.Raise 9, Err.Source, "Color Index is out of bounds. The index must be between 1 and 59.", Err.HelpFile, Err.HelpContext
    End If

End Property
Private Sub Form_Load()
    Dim Top             As Long
    Dim Left            As Long
    Dim Red             As Integer
    Dim Grn             As Integer
    Dim Blu             As Integer
    
    With Me
        '   Load the Excel ColorIndex values
        Call LoadColors
        '   Now create an array of PictureBoxes to
        '   serve as the individual color picker items
        k = 0
        For i = 1 To 7
            For j = 1 To 8
                k = k + 1
                Load Me.P1(k)
                With Me.P1(k)
                    .Height = 255
                    .Width = 255
                    Top = (.Height * i) - 128
                    Left = (.Width * j) - 128
                    .Top = Top
                    .Left = Left
                    .BackColor = xlColors(k)
                    .ToolTipText = xlColors(k)
                    .BorderStyle = 1
                    .Visible = True
                End With
            Next
        Next i
        '   Give the form an initial value based on the local property
        .P1(ColorIndex).BorderStyle = 0
        m_PrevIndex = ColorIndex
    End With
End Sub

Private Sub LoadColors()
    '   This is the order Excel assigned to the ColorIndex value,
    '   so we will repeat it here...
    ReDim xlColors(1 To 56)
    xlColors(1) = &H0&
    xlColors(2) = &HFFFFFF
    xlColors(3) = &HFF&
    xlColors(4) = &HFF00&
    xlColors(5) = &HFF0000
    xlColors(6) = &HFFFF&
    xlColors(7) = &HFF00FF
    xlColors(8) = &HFFFF00
    xlColors(9) = &H80&
    xlColors(10) = &H8000&
    xlColors(11) = &H800000
    xlColors(12) = &H8080&
    xlColors(13) = &H800080
    xlColors(14) = &H808000
    xlColors(15) = &HC0C0C0
    xlColors(16) = &H808080
    xlColors(17) = &HFF9999
    xlColors(18) = &H663399
    xlColors(19) = &HCCFFFF
    xlColors(20) = &HFFFFCC
    xlColors(21) = &H660066
    xlColors(22) = &H8080FF
    xlColors(23) = &HCC6600
    xlColors(24) = &HFFCCCC
    xlColors(25) = &H800000
    xlColors(26) = &HFF00FF
    xlColors(27) = &HFFFF&
    xlColors(28) = &HFFFF00
    xlColors(29) = &H800080
    xlColors(30) = &H80&
    xlColors(31) = &H808000
    xlColors(32) = &HFF0000
    xlColors(33) = &HFFCC00
    xlColors(34) = &HFFFFCC
    xlColors(35) = &HCCFFCC
    xlColors(36) = &H99FFFF
    xlColors(37) = &HFFCC99
    xlColors(38) = &HCC99FF
    xlColors(39) = &HFF99CC
    xlColors(40) = &H99CCFF
    xlColors(41) = &HFF6633
    xlColors(42) = &HCCCC33
    xlColors(43) = &HCC99&
    xlColors(44) = &HCCFF&
    xlColors(45) = &H99FF&
    xlColors(46) = &H66FF&
    xlColors(47) = &H996666
    xlColors(48) = &H969696
    xlColors(49) = &H663300
    xlColors(50) = &H669933
    xlColors(51) = &H3300&
    xlColors(52) = &H3333&
    xlColors(53) = &H3399&
    xlColors(54) = &H663399
    xlColors(55) = &H993333
    xlColors(56) = &H333333
End Sub

Private Sub Ok_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub P1_Click(Index As Integer)
'    If Index = 0 Then Index = 1
    Me.P1.Item(m_PrevIndex).BorderStyle = 1
    Me.P1.Item(Index).BorderStyle = 0
    m_PrevIndex = Index
    ExampleForm.ColorValue.Text = Index
End Sub


