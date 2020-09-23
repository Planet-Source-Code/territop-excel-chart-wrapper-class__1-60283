VERSION 5.00
Begin VB.Form ExampleForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel Wrapper Class Example"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   Icon            =   "ExcelOLEForm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   240
      Top             =   4440
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   27
      ToolTipText     =   "Close Example"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Options 
      Caption         =   "Options >>"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Settings for selected properties"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton LoadPlot 
      Caption         =   "Plot"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      ToolTipText     =   "Plot Test Data"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart"
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   7335
      Begin VB.CommandButton LoadPallete 
         Caption         =   "..."
         Height          =   255
         Left            =   1650
         TabIndex        =   26
         Top             =   1475
         Width           =   255
      End
      Begin VB.TextBox ColorValue 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Text            =   "ColorValue"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Interior 
         Caption         =   "Interior"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         ToolTipText     =   "Chart Interior Color"
         Top             =   1120
         Width           =   855
      End
      Begin VB.OptionButton Boarder 
         Caption         =   "Border "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Chart Border Color"
         Top             =   1120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox YAxisTitle 
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Text            =   "Concentration (mg/ml)"
         Top             =   1440
         Width           =   1800
      End
      Begin VB.TextBox XAxisTitle 
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Text            =   "Time (min)"
         Top             =   985
         Width           =   1800
      End
      Begin VB.TextBox ChartTitle 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Text            =   "Time Course"
         Top             =   525
         Width           =   1800
      End
      Begin VB.ComboBox ChartStyles 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "ChartType"
         ToolTipText     =   "Select a Chart Type"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox HasTitle 
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Enable Chart Title"
         Top             =   525
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox HasXTitle 
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Enable X Axis Title"
         Top             =   985
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox HasYTitle 
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Enable Y Axis Title"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Frame Frame2 
         Height          =   1400
         Left            =   5160
         TabIndex        =   20
         Top             =   400
         Width           =   2035
         Begin VB.CheckBox MinorGrid 
            Caption         =   "Minor Gridlines"
            Height          =   255
            Left            =   200
            TabIndex        =   24
            ToolTipText     =   "Chart Minor Gridlines"
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox MajorGrid 
            Caption         =   "Major Gridlines"
            Height          =   255
            Left            =   200
            TabIndex        =   23
            ToolTipText     =   "Chart Major Gridlines"
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton YAxis 
            Caption         =   "Y-Axis"
            Height          =   255
            Left            =   1080
            TabIndex        =   22
            ToolTipText     =   "Y Axis Gridlines"
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton XAxis 
            Caption         =   "X-Axis"
            Height          =   255
            Left            =   200
            TabIndex        =   21
            ToolTipText     =   "X Axis Gridlines"
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Gridlines:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Titles:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblYAxis 
         Caption         =   "Y-Title:"
         Height          =   255
         Left            =   2535
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblXAxis 
         Caption         =   "X-Title:"
         Height          =   255
         Left            =   2535
         TabIndex        =   11
         Top             =   985
         Width           =   495
      End
      Begin VB.Label lblTitle 
         Caption         =   "Chart:"
         Height          =   255
         Left            =   2535
         TabIndex        =   7
         Top             =   525
         Width           =   495
      End
      Begin VB.Label lblType 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      Caption         =   "To edit the chart, double click on the chart object to edit it directly, or click ""Options"" for selected items."
      Height          =   375
      Left            =   1440
      TabIndex        =   28
      Top             =   5040
      Width           =   3735
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Chart.8"
      Height          =   4695
      Left            =   210
      OleObjectBlob   =   "ExcelOLEForm2.frx":57E2
      TabIndex        =   1
      Top             =   210
      Width           =   7395
   End
End
Attribute VB_Name = "ExampleForm"
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
'       30Apr05 - Bug fixes with the test harness, OLEClass Untounched.
'
'   Force Declarations
Option Explicit
Private CExcelPlot              As CExcelPlot   'Excel OLE2 Wrapper Class
Private i                       As Integer      'Loop Counter
Private m_AxisClicked           As Boolean      'Axis Clicked Flag
Private m_BoarderColor          As Long         'Border Color of the Chart
Private m_InteriorColor         As Long         'Chart Interior Color
Private m_LoadColorPallete      As Boolean      'Color Pallete Loaded Flag
Private m_PlotClicked           As Boolean      'Plot Clicked Flag
Private m_Style(0 To 74)        As Long         'Chart Style
Private m_XMajorGrid            As Long         'X Axis Major Gridlines
Private m_XMinorGrid            As Long         'X Axis Minor Gridlines
Private m_YMajorGrid            As Long         'Y Axis Major Gridlines
Private m_YMinorGrid            As Long         'Y Axis Minor Gridlines
Private XArr(1 To 15, 1 To 1)   As Double       'X Data Array
Private XErr(1 To 15, 1 To 1)   As Double       'X Error Data Array
Private YArr(1 To 15, 1 To 2)   As Double       'Y Data Array
Private YErr(1 To 15, 1 To 2)   As Double       'Y Error Data Array

'Private Sub XLView_Click()
'    OLE1.DoVerb vbOLEOpen
'End Sub

Private Sub Boarder_Click()
    With Me
        If .Boarder.Value = True Then
            .ColorValue.Text = m_BoarderColor
        Else
            .ColorValue.Text = m_InteriorColor
        End If
    End With
End Sub

Private Sub Cancel_Click()
    If MsgBox("Close the Example Form?", vbYesNo, "Excel OLE Example") = vbYes Then
        Form_Terminate
    End If
End Sub

Private Sub ChartStyles_Change()
    With Me
        .ChartStyle = m_Style(Me.ChartStyles.ListIndex)
        '   Now plot the data with the changes...
        LoadPlot_Click
    End With
End Sub

Private Sub ColorValue_Change()
    With Me
        If IsNumeric(ColorValue.Text) And ((ColorValue.Text >= 1) And (ColorValue.Text < 57)) Then
            If .Boarder.Value = True Then
                m_BoarderColor = CLng(ColorValue.Text)
            Else
                m_InteriorColor = CLng(ColorValue.Text)
            End If
        Else
            MsgBox "Please enter a valid number.", vbExclamation, "Excel OLE Example"
        End If
    End With
End Sub

Private Sub Form_Click()
    '   If the user DblClicks the OLE2 object to edit,
    '   we need to provide a method to return back which
    '   is easy to use and doesn't hang the application....
    SendKeys "{Esc}"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Load frmSplash
    DoEvents
    frmSplash.Show
    '   Initialize the example form and setup the Wrapper
    Set CExcelPlot = New CExcelPlot
    frmSplash.Hide
    Unload frmSplash
    With Me.ChartStyles
        .AddItem "3DArea"
        .AddItem "3DAreaStacked"
        .AddItem "3DAreaStacked100"
        .AddItem "3DBarClustered"
        .AddItem "3DBarStacked"
        .AddItem "3DBarStacked100"
        .AddItem "3DColumn"
        .AddItem "3DColumnClustered"
        .AddItem "3DColumnStacked"
        .AddItem "3DColumnStacked100"
        .AddItem "3DLine"
        .AddItem "3DPie"
        .AddItem "3DPieExploded"
        .AddItem "3DSurface"
        .AddItem "Area"
        .AddItem "AreaStacked"
        .AddItem "AreaStacked100"
        .AddItem "BarClustered"
        .AddItem "BarOfPie"
        .AddItem "BarStacked"
        .AddItem "BarStacked100"
        .AddItem "Bubble"
        .AddItem "Bubble3DEffect"
        .AddItem "ColumnClustered"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked100"
        .AddItem "Combination"
        .AddItem "ConeBarClustered"
        .AddItem "ConeBarStacked"
        .AddItem "ConeBarStacked100"
        .AddItem "ConeCol"
        .AddItem "ConeColClustered"
        .AddItem "ConeColStacked"
        .AddItem "ConeColStacked100"
        .AddItem "CylinderBarClustered"
        .AddItem "CylinderBarStacked"
        .AddItem "CylinderBarStacked100"
        .AddItem "CylinderCol"
        .AddItem "CylinderColClustered"
        .AddItem "CylinderColStacked"
        .AddItem "CylinderColStacked100"
        .AddItem "Doughnut"
        .AddItem "DoughnutExploded"
        .AddItem "Line"
        .AddItem "LineMarkers"
        .AddItem "LineMarkersStacked"
        .AddItem "LineMarkersStacked100"
        .AddItem "LineStacked"
        .AddItem "LineStacked100"
        .AddItem "Pie"
        .AddItem "PieExploded"
        .AddItem "PieOfPie"
        .AddItem "PyramidBarClustered"
        .AddItem "PyramidBarStacked"
        .AddItem "PyramidBarStacked100"
        .AddItem "PyramidCol"
        .AddItem "PyramidColClustered"
        .AddItem "PyramidColStacked"
        .AddItem "PyramidColStacked100"
        .AddItem "Radar"
        .AddItem "RadarFilled"
        .AddItem "RadarMarkers"
        .AddItem "StockHLC"
        .AddItem "StockOHLC"
        .AddItem "StockVHLC"
        .AddItem "StockVOHLC"
        .AddItem "Surface"
        .AddItem "SurfaceTopView"
        .AddItem "SurfaceTopViewWireframe"
        .AddItem "SurfaceWireframe"
        .AddItem "XYScatter"
        .AddItem "XYScatterLines"
        .AddItem "XYScatterLinesNoMarkers"
        .AddItem "XYScatterSmooth"
        .AddItem "XYScatterSmoothNoMarkers"
        .ListIndex = 70
    End With
    '   Fill the local Array of types from the  plotting class
    '   so we can refer to them via indexes....
    m_Style(0) = xl3DArea
    m_Style(1) = xl3DAreaStacked
    m_Style(2) = xl3DAreaStacked100
    m_Style(3) = xl3DBarClustered
    m_Style(4) = xl3DBarStacked
    m_Style(5) = xl3DBarStacked100
    m_Style(6) = xl3DColumn
    m_Style(7) = xl3DColumnClustered
    m_Style(8) = xl3DColumnStacked
    m_Style(9) = xl3DColumnStacked100
    m_Style(10) = xl3DLine
    m_Style(11) = xl3DPie
    m_Style(12) = xl3DPieExploded
    m_Style(13) = xl3DSurface
    m_Style(14) = xlArea
    m_Style(15) = xlAreaStacked
    m_Style(16) = xlAreaStacked100
    m_Style(17) = xlBarClustered
    m_Style(18) = xlBarOfPie
    m_Style(19) = xlBarStacked
    m_Style(20) = xlBarStacked100
    m_Style(21) = xlBubble
    m_Style(22) = xlBubble3DEffect
    m_Style(23) = xlColumnClustered
    m_Style(24) = xlColumnStacked
    m_Style(25) = xlColumnStacked100
    m_Style(26) = xlCombination
    m_Style(27) = xlConeBarClustered
    m_Style(28) = xlConeBarStacked
    m_Style(29) = xlConeBarStacked100
    m_Style(30) = xlConeCol
    m_Style(31) = xlConeColClustered
    m_Style(32) = xlConeColStacked
    m_Style(33) = xlConeColStacked100
    m_Style(34) = xlCylinderBarClustered
    m_Style(35) = xlCylinderBarStacked
    m_Style(36) = xlCylinderBarStacked100
    m_Style(37) = xlCylinderCol
    m_Style(38) = xlCylinderColClustered
    m_Style(39) = xlCylinderColStacked
    m_Style(40) = xlCylinderColStacked100
    m_Style(41) = xlDoughnut
    m_Style(42) = xlDoughnutExploded
    m_Style(43) = xlLine
    m_Style(44) = xlLineMarkers
    m_Style(45) = xlLineMarkersStacked
    m_Style(46) = xlLineMarkersStacked100
    m_Style(47) = xlLineStacked
    m_Style(48) = xlLineStacked100
    m_Style(49) = xlPie
    m_Style(50) = xlPieExploded
    m_Style(51) = xlPieOfPie
    m_Style(52) = xlPyramidBarClustered
    m_Style(53) = xlPyramidBarStacked
    m_Style(54) = xlPyramidBarStacked100
    m_Style(55) = xlPyramidCol
    m_Style(56) = xlPyramidColClustered
    m_Style(57) = xlPyramidColStacked
    m_Style(58) = xlPyramidColStacked100
    m_Style(59) = xlRadar
    m_Style(60) = xlRadarFilled
    m_Style(61) = xlRadarMarkers
    m_Style(62) = xlStockHLC
    m_Style(63) = xlStockOHLC
    m_Style(64) = xlStockVHLC
    m_Style(65) = xlStockVOHLC
    m_Style(66) = xlSurface
    m_Style(67) = xlSurfaceTopView
    m_Style(68) = xlSurfaceTopViewWireframe
    m_Style(69) = xlSurfaceWireframe
    m_Style(70) = xlXYScatter
    m_Style(71) = xlXYScatterLines
    m_Style(72) = xlXYScatterLinesNoMarkers
    m_Style(73) = xlXYScatterSmooth
    m_Style(74) = xlXYScatterSmoothNoMarkers
    
    '   Set the initial form conditions
    With Me
        .Options.Caption = "Options >>"
        .Height = 6040
        .Boarder.Value = True
        .XAxis.Value = True
        m_XMajorGrid = vbUnchecked
        m_YMajorGrid = vbUnchecked
        m_XMinorGrid = vbUnchecked
        m_XMinorGrid = vbUnchecked
        m_BoarderColor = 2          'xlWhite
        m_InteriorColor = 19        'xlCream
        '   Init the color value value for the boarder
        .ColorValue.Text = m_BoarderColor
    End With
    '   Now plot the data with the changes...
    LoadPlot_Click
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        '   If the user DblClicks the OLE2 object to edit,
        '   we need to provide a method to return back which
        '   is easy to use and doesn't hang the application....
        SendKeys "{Esc}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   Going somewhere?
    If MsgBox("Close the Example Form?", vbYesNo, "Excel OLE Example") = vbYes Then
        Cancel = 0
        Form_Terminate
    Else
        Cancel = 1
    End If
End Sub

Private Sub Form_Terminate()
    Me.Hide
    '   Free up memory before we go...
    Set CExcelPlot.OLEContainer = Nothing
    Set CExcelPlot = Nothing
    Set ExampleForm = Nothing
    End
End Sub

Private Sub HasTitle_Click()
    With Me
        If .HasTitle.Value = vbUnchecked Then
            .ChartTitle.Enabled = False
            .lblTitle.Enabled = False
        Else
            .ChartTitle.Enabled = True
            .lblTitle.Enabled = True
        End If
    End With
    '   Now plot the data with the changes...
    LoadPlot_Click
End Sub

Private Sub HasXTitle_Click()
    With Me
        If .HasXTitle.Value = vbUnchecked Then
            .XAxisTitle.Enabled = False
            .lblXAxis.Enabled = False
        Else
            .XAxisTitle.Enabled = True
            .lblXAxis.Enabled = True
        End If
    End With
    '   Now plot the data with the changes...
    LoadPlot_Click
End Sub

Private Sub HasYTitle_Click()
    With Me
        If .HasYTitle.Value = vbUnchecked Then
            .YAxisTitle.Enabled = False
            .lblYAxis.Enabled = False
        Else
            .YAxisTitle.Enabled = True
            .lblYAxis.Enabled = True
        End If
    End With
End Sub

Private Sub Interior_Click()
    '   Get the old values to disaplay to the user
    With Me
        If .Boarder.Value = True Then
            .ColorValue.Text = m_BoarderColor
        Else
            .ColorValue.Text = m_InteriorColor
        End If
    End With
End Sub

Private Sub LoadPallete_Click()
    '   Load the Excel Color Pallete
    ColorPallete.ColorIndex = CLng(Me.ColorValue.Text)
    Load ColorPallete
    ColorPallete.Show
End Sub

Private Sub LoadPlot_Click()
    '   Plot some data to illustrate the Wrapper Class
    Dim iCol            As Integer
    Dim iRow            As Integer
    
    Screen.MousePointer = vbHourglass
    Set CExcelPlot.OLEContainer = Me.OLE1
    
    For iCol = 1 To 2
        For iRow = 1 To 15
            XArr(iRow, 1) = iRow
            '   Create some random noise in X
            XErr(iRow, 1) = iRow * Rnd() * 0.1
            YArr(iRow, iCol) = iCol * 80 * Exp(-iRow * 0.1) + 60 * Exp(-iRow * 0.15) + (-170 * Exp(-iRow * 0.5))
            '   Create some random noise in Y
            YErr(iRow, iCol) = YArr(iRow, iCol) * Rnd() * 0.1
        Next iRow
    Next iCol
    With CExcelPlot
        .ChartTitle = Me.ChartTitle
        .HasChartTitle = Me.HasTitle
        .XAxisTitle = Me.XAxisTitle
        .HasXAxisTitle = Me.HasXTitle
        .YAxisTitle = Me.YAxisTitle
        .HasYAxisTitle = Me.HasYTitle
        .ChartStyle = m_Style(Me.ChartStyles.ListIndex)
        .InteriorColorIndex = m_InteriorColor
        .ChartColorIndex = m_BoarderColor
        .HasXMajorGrid = m_XMajorGrid
        .HasXMinorGrid = m_XMinorGrid
        .HasYMajorGrid = m_YMajorGrid
        .HasYMinorGrid = m_YMinorGrid
        .Plot XArr, YArr, XErr, YErr
    End With
    m_PlotClicked = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub MajorGrid_Click()
    '   Set the old values to disaplay to the user
    With Me
        If (m_AxisClicked = False) Then
            If (.XAxis.Value = True) Then
                m_XMajorGrid = .MajorGrid.Value
                m_XMinorGrid = .MinorGrid.Value
            Else
                m_YMajorGrid = .MajorGrid.Value
                m_YMinorGrid = .MinorGrid.Value
            End If
        End If
    End With
    '   Now plot the data with the changes...
    LoadPlot_Click
End Sub

Private Sub MinorGrid_Click()
    '   Set the old values to disaplay to the user
    With Me
        If (m_AxisClicked = False) Then
            If (.XAxis.Value = True) Then
                m_XMajorGrid = .MajorGrid.Value
                m_XMinorGrid = .MinorGrid.Value
            Else
                m_YMajorGrid = .MajorGrid.Value
                m_YMinorGrid = .MinorGrid.Value
            End If
        End If
    End With
    '   Now plot the data with the changes...
    LoadPlot_Click
End Sub

Private Sub Options_Click()
    With Me
        If .Options.Caption = "Options >>" Then
            .Options.Caption = "<< Options"
            .Height = 8140
        Else
            .Options.Caption = "Options >>"
            .Height = 6040
            '   Call the events which normally will store
            '   local copies of the variables
            ColorValue_Change
        End If
    End With
End Sub

'Private Sub Timer1_Timer()
'   '   This sub will cursor through all of the color styles, if one
'   '   wants to use it feel free, but i used this to debug the colorindexes
'    Dim iCol            As Integer
'    Dim iRow            As Integer
'
'    If m_PlotClicked = True Then
'        m_BoarderColor = m_BoarderColor + 1 Mod 56
'        With CExcelPlot
'            .ChartTitle = Me.ChartTitle
'            .XAxisTitle = Me.XAxisTitle
'            .YAxisTitle = Me.YAxisTitle
'            .ChartStyle = m_Style(Me.ChartStyles.ListIndex)
'            .InteriorColor = m_InteriorColor
'            .ChartColor = m_BoarderColor
'            .HasXMajorGrid = m_XMajorGrid
'            .HasXMinorGrid = m_XMinorGrid
'            .HasYMajorGrid = m_YMajorGrid
'            .HasYMinorGrid = m_YMinorGrid
'            .Plot XArr, YArr
'        End With
'    End If
'End Sub

Private Sub XAxis_Click()
    '   Get the old values to disaplay to the user
    With Me
        m_AxisClicked = True
        If .XAxis.Value = True Then
            .MajorGrid.Value = m_XMajorGrid
            .MinorGrid.Value = m_XMinorGrid
        Else
            .MajorGrid.Value = m_YMajorGrid
            .MinorGrid.Value = m_YMinorGrid
        End If
        m_AxisClicked = False
    End With
End Sub

Private Sub YAxis_Click()
    '   Get the old values to disaplay to the user
    With Me
        m_AxisClicked = True
        If .XAxis.Value = True Then
            .MajorGrid.Value = m_XMajorGrid
            .MinorGrid.Value = m_XMinorGrid
        Else
            .MajorGrid.Value = m_YMajorGrid
            .MinorGrid.Value = m_YMinorGrid
        End If
        m_AxisClicked = False
    End With
End Sub

                    
