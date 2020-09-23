VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      Begin VB.Timer Timer1 
         Left            =   4680
         Top             =   840
      End
      Begin VB.Label Label2 
         Caption         =   "by Paul R. Territo, Ph.D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label LoadLabel 
         Caption         =   "Initilizing: Excel OL2 Connection...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   3255
      End
      Begin VB.Image imgLogo 
         Height          =   615
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Excel OLE2 Wrapper Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3720
      End
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    '   Start the timer and permit events
    Timer1.Enabled = True
    Timer1.Interval = 100
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '   Make sure the timer is stopped
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    '   This timer is simply for flashing a label....and does not
    '   add to the loading times....
    DoEvents
    '   Flash the label, just in case this take a long time...
    Me.LoadLabel.Visible = -(Me.LoadLabel.Visible + 1 Mod 2)
End Sub
