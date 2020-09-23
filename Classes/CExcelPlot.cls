VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExcelPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Excel Wrapper Class based on OLE2 Embedding "
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'   API Declarations
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'   Local Properties
Private m_ChartColorIndex   As Long         'Chart Color Index
Private m_ChartColorVal     As Long         'Chart Color Long Color Value
Private m_ChartTitleSize    As Long         'Size of the Font for the Title
Private m_ChartStyle        As ChartType    'Chart Style
Private m_ChartTitle        As String       'Chart Title
Private m_Colors()          As Long         'Excel Color Table
Private m_DataLayout        As Orienation   'Layout of the Arrays Passed
Private m_HasChartTitle     As Boolean      'Chart Title
Private m_HasPlotLegend     As Boolean      'Plot Legend?
Private m_HasXAxisTitle     As Boolean      'X Axis Title?
Private m_HasXMajorGrid     As Boolean      'X Axis Major Gridlines?
Private m_HasXMinorGrid     As Boolean      'X Axis Minor Gridlines?
Private m_HasYAxisTitle     As Boolean      'Y Axis Title?
Private m_HasYMajorGrid     As Boolean      'Y Axis Major Gridlines?
Private m_HasYMinorGrid     As Boolean      'Y Axis Minor Gridlines?
Private m_InteriorColorIndex As Long        'Plot Interior Color Index
Private m_InteriorColorVal  As Long         'Plot Interior Long Color Value
Private m_ReversePlotOrder  As Boolean      'Reverse the order of the plotted data
Private m_XAxisMajorUnitsAuto As Boolean    'Major Units of the X Axis is Automatically Determined
Private m_XAxisMaxScaleAuto As Boolean      'Maximum Value of the X Axis is Automatically Determined
Private m_XAxisMinorUnitsAuto As Boolean    'Minor Units of the X Axis is Automatically Determined
Private m_XAxisMinScaleAuto As Boolean      'Minimum Value of the X Axis is Automatically Determined
Private m_XAxisScaleType    As ScaleMode    'Scale Mode of the X Axis (Linear, Log10)
Private m_XAxisTitle        As String       'X Axis Title
Private m_XLabelSize        As Long         'X Axis Label Size
Private m_XTickLabelSize    As Long         'X Axis Tick Label Size
Private m_YAxisMajorUnitsAuto As Boolean    'Major Units of the Y Axis is Automatically Determined
Private m_YAxisMaxScaleAuto As Boolean      'Maximum Value of the Y Axis is Automatically Determined
Private m_YAxisMinorUnitsAuto As Boolean    'Minor Units of the Y Axis is Automatically Determined
Private m_YAxisMinScaleAuto As Boolean      'Minimum Value of the Y Axis is Automatically Determined
Private m_YAxisScaleType    As ScaleMode    'Scale Mode of the Y Axis (Linear, Log10)
Private m_YAxisTitle        As String       'Y Axis Title
Private m_YLabelSize        As Long         'Y Axis Label Size
Private m_YTickLabelSize    As Long         'Y Axis Tick Label Size
Private oChart              As Chart        'Object Reference to Excel Chart
Private oOLE                As OLE          'OLE2 Object Container
Private oSheet              As Worksheet    'Object Reference to Excel Worksheet
Private oExcel              As Excel.Application    'Excel Instance
'   Public Enumerations
Public Enum ChartType
    xl3DArea = -4098                        '3D Area Plot
    xl3DAreaStacked = 78                    '3D Area Plot - Stacked
    xl3DAreaStacked100 = 79                 '3D Area Plot - Stacked Scaled to 100%
    xl3DBarClustered = 60                   '3D Horizontal Bar Plot - Clustered
    xl3DBarStacked = 61                     '3D Horizontal Bar Plot - Stacked
    xl3DBarStacked100 = 62                  '3D Bar - Stacked Scaled to 100%
    xl3DColumn = -4100                      '3D Vertical Bar Plot
    xl3DColumnClustered = 54                '3D Vertical Bar Plot - Clustered
    xl3DColumnStacked = 55                  '3D Vertical Bar Plot - Stacked
    xl3DColumnStacked100 = 56               '3D Vertical Bar Plot - Stacked Scaled to 100%
    xl3DLine = -4101                        '3D Line Plot
    xl3DPie = -4102                         '3D Pie Plot
    xl3DPieExploded = 70                    '3D Pie Plot - Exploded
    xl3DSurface = -4103                     '3D Surface Plot
    xlArea = 1                              'Area Plot
    xlAreaStacked = 76                      'Area Plot - Stacked
    xlAreaStacked100 = 77                   'Area Plot - Stacked Scaled to 100%
    xlBarClustered = 57                     'Horizontal Bar Plot
    xlBarOfPie = 71                         'Horizontal Bar & Pie Plot Combined
    xlBarStacked = 58                       'Horizontal Bar Plot - Stacked
    xlBarStacked100 = 59                    'Horizontal Bar Plot - Stacked Scaled to 100%
    xlBubble = 15                           'Bubble Plot
    xlBubble3DEffect = 87                   '3D Bubble Plot
    xlColumnClustered = 51                  'Verical Bar Plot
    xlColumnStacked = 52                    'Verical Bar Plot - Stacked
    xlColumnStacked100 = 53                 'Verical Bar Plot - Stacked Scaled to 100%
    xlCombination = -4111                   'Any Two Plot Styles - Except 3D
    xlConeBarClustered = 102                'Cone Styled Bar Plot Clustered
    xlConeBarStacked = 103                  'Cone Styled Bar Plot - Stacked
    xlConeBarStacked100 = 104               'Cone Styled Bar Plot - Stacked Scaled to 100%
    xlConeCol = 105                         'Cone Column Plot
    xlConeColClustered = 99                 'Cone Column Plot - Clusterd
    xlConeColStacked = 100                  'Cone Column Plot - Stacked
    xlConeColStacked100 = 101               'Cone Column Plot - Stacked Scaled to 100%
    xlCylinderBarClustered = 95             'Cylinder Bar - Clustered
    xlCylinderBarStacked = 96               'Cylinder Bar - Stacked
    xlCylinderBarStacked100 = 97            'Cylinder Bar - Stacked Scaled to 100%
    xlCylinderCol = 98                      'Vertical Cylinder Plot
    xlCylinderColClustered = 92             'Vertical Cylinder Plot - Clustered
    xlCylinderColStacked = 93               'Vertical Cylinder Plot -  Stacked
    xlCylinderColStacked100 = 94            'Vertical Cylinder Plot - Stacked Scaled to 100%
    xlDoughnut = -4120                      'Pie Plot Minus the Center
    xlDoughnutExploded = 80                 'Pie Plot Minus the Center - Exploded
    xlLine = 4                              'Line Plot
    xlLineMarkers = 65                      'Line Plot with Markers
    xlLineMarkersStacked = 66               'Line Plot with Markers - Stacked
    xlLineMarkersStacked100 = 67            'Line Plot with Markers - Stacked Scaled to 100%
    xlLineStacked = 63                      'Line Plot without Markers - Stacked
    xlLineStacked100 = 64                   'Line Plot without Markers - Stacked Scaled to 100%
    xlPie = 5                               'Pie Plot
    xlPieExploded = 69                      'Pie Plot - Exploded
    xlPieOfPie = 68                         'Pie of Pie Plot
    xlPyramidBarClustered = 109             'Horizontal Pyramidal Bar - Clustered
    xlPyramidBarStacked = 110               'Horizontal Pyramidal Bar - Stacked
    xlPyramidBarStacked100 = 111            'Horizontal Pyramidal Bar - Stacked Scaled to 100%
    xlPyramidCol = 112                      'Vertical Pyramidal Bar -
    xlPyramidColClustered = 106             'Vertical Pyramidal Bar - Clustered
    xlPyramidColStacked = 107               'Vertical Pyramidal Bar - Stacked
    xlPyramidColStacked100 = 108            'Vertical Pyramidal Bar - Stacked Scaled to 100%
    xlRadar = -4151                         'Polar Coordinate Plot
    xlRadarFilled = 82                      'Polar Coordinate Plot - Filled
    xlRadarMarkers = 81                     'Polar Coordinate Plot - Filled with Markers
    xlStockHLC = 88                         'Stock Plot with Hi/Lo in Columns
    xlStockOHLC = 89                        'Stock Plot with Hi/Lo in Columns / Horizontal
    xlStockVHLC = 90                        'Stock Plot with Hi/Lo in Vertical / Horizontal
    xlStockVOHLC = 91                       'Stock Plot with Hi/Lo in Vertical
    xlSurface = 83                          'Surface Plot
    xlSurfaceTopView = 85                   'Surface Top View Plot
    xlSurfaceTopViewWireframe = 86          'Surface Top View Wireframe Plot
    xlSurfaceWireframe = 84                 'Surface Wireframe Plot
    xlXYScatter = -4169                     'XY Scatter Plot
    xlXYScatterLines = 74                   'XY Scatter Plot with Lines
    xlXYScatterLinesNoMarkers = 75          'XY Scatter Plot with Lines without Markers
    xlXYScatterSmooth = 72                  'XY Scatter Plot - Smooth
    xlXYScatterSmoothNoMarkers = 73         'XY Scatter Plot - Smooth without Markers
End Enum

Public Enum Orienation
    xlRows = 1                              'Data Arranged in Rows
    xlColumns = 2                           'Data Arranged in Columns
End Enum

Public Enum ScaleMode
    xlLinear = -4132                        'Linear Scaling of the Axis
    xlLogarithmic = -4133                   'Log10 Scaling of the Axis
End Enum

Public Property Get ChartColorIndex() As Long
    ChartColorIndex = m_ChartColorIndex
End Property

Public Property Let ChartColorIndex(Color As Long)
    '   Check to see if we are within range
    If (Color >= 1) And (Color < 57) Then
        m_ChartColorIndex = Color
    Else
        '   We are not so tell the user
        Err.Raise 9, Err.Source, "Color Index is out of bounds. The index must be between 1 and 59.", Err.HelpFile, Err.HelpContext
    End If
End Property

Public Property Get ChartColorVal() As Long
    ChartColorVal = m_ChartColorVal
End Property

Public Property Let ChartColorVal(Color As Long)
    m_ChartColorVal = Color
End Property

Public Property Get ChartStyle() As ChartType
    ChartStyle = m_ChartStyle
End Property

Public Property Let ChartStyle(Style As ChartType)
    m_ChartStyle = Style
End Property

Public Property Get ChartTitle() As String
    ChartTitle = m_ChartTitle
End Property

Public Property Let ChartTitle(title As String)
    m_ChartTitle = title
End Property

Public Property Get ChartTitleSize() As String
    ChartTitleSize = m_ChartTitleSize
End Property

Public Property Let ChartTitleSize(LabelSize As String)
    m_ChartTitleSize = LabelSize
End Property

Private Sub Class_Initialize()
    '   Set the mouse pointer to Hourglass in case this takes a while to load
    Screen.MousePointer = vbHourglass
    '   Speed up plotting of the Excel Wrapper Class by pre-loading a
    '   Private instance of The Excel Object
    Set oExcel = New Excel.Application
    '   Initialize the properties
    m_ChartColorIndex = 2                   'Chart Color Index - xlWhite
    m_ChartColorVal = -1                    'Chart Color Value - Unset (-1)
    m_ChartTitleSize = 12                   'Chart Title Font Size
    m_ChartStyle = xlXYScatter              'Init Style of Plot
    m_ChartTitle = "Chart Title"            'Chart Title
    m_DataLayout = xlColumns                'Data Layout as Columns
    m_HasChartTitle = True                  'Chart Title
    m_HasPlotLegend = False                 'Plot Legend?
    m_HasXAxisTitle = True                  'X Axis Title?
    m_HasXMajorGrid = False                 'X Axis Major Gridlines?
    m_HasXMinorGrid = False                 'X Axis Minor Gridlines?
    m_HasYAxisTitle = True                  'Y Axis Title?
    m_HasYMajorGrid = False                 'Y Axis Major Gridlines?
    m_HasYMinorGrid = False                 'Y Axis Minor Gridlines?
    m_InteriorColorIndex = 19               'Plot Interior Color Index - xlCream
    m_InteriorColorVal = -1                 'Plot Interior Color Value - Unset (-1)
    m_ReversePlotOrder = False              'Reverse the order of the plotted data
    m_XAxisMajorUnitsAuto = True            'Major Units of the X Axis is Automatically Determined
    m_XAxisMaxScaleAuto = True              'Maximum Value of the X Axis is Automatically Determined
    m_XAxisMinorUnitsAuto = True            'Minor Units of the X Axis is Automatically Determined
    m_XAxisMinScaleAuto = True              'Minimum Value of the X Axis is Automatically Determined
    m_XAxisScaleType = xlLinear             'Scale Mode of the X Axis (Linear, Log10)
    m_XAxisTitle = "X Axis"                 'X Axis Title
    m_XLabelSize = 10                       'X Axis Label Size
    m_XTickLabelSize = 10                   'X Axis Tick Label Size
    m_YAxisMajorUnitsAuto = True            'Major Units of the Y Axis is Automatically Determined
    m_YAxisMaxScaleAuto = True              'Maximum Value of the Y Axis is Automatically Determined
    m_YAxisMinorUnitsAuto = True            'Minor Units of the Y Axis is Automatically Determined
    m_YAxisMinScaleAuto = True              'Minimum Value of the Y Axis is Automatically Determined
    m_YAxisScaleType = xlLinear             'Scale Mode of the Y Axis (Linear, Log10)
    m_YAxisTitle = "Y Axis"                 'Y Axis Title
    m_YLabelSize = 10                       'Y Axis Label Size
    m_YTickLabelSize = 10                   'Y Axis Tick Label Size
    '
    '   This is the value Excel assigned to the ColorIndex,
    '   so we will repeat it here...
    '
    '   Note: if one want to use Hex color value directly, then the
    '   subsequent call passes the Hex value to the .Color property
    '   of the specified object (i.e. ChartArea, PlotArea)....
    ReDim m_Colors(1 To 56)
    m_Colors(1) = &H0&
    m_Colors(2) = &HFFFFFF
    m_Colors(3) = &HFF&
    m_Colors(4) = &HFF00&
    m_Colors(5) = &HFF0000
    m_Colors(6) = &HFFFF&
    m_Colors(7) = &HFF00FF
    m_Colors(8) = &HFFFF00
    m_Colors(9) = &H80&
    m_Colors(10) = &H8000&
    m_Colors(11) = &H800000
    m_Colors(12) = &H8080&
    m_Colors(13) = &H800080
    m_Colors(14) = &H808000
    m_Colors(15) = &HC0C0C0
    m_Colors(16) = &H808080
    m_Colors(17) = &HFF9999
    m_Colors(18) = &H663399
    m_Colors(19) = &HCCFFFF
    m_Colors(20) = &HFFFFCC
    m_Colors(21) = &H660066
    m_Colors(22) = &H8080FF
    m_Colors(23) = &HCC6600
    m_Colors(24) = &HFFCCCC
    m_Colors(25) = &H800000
    m_Colors(26) = &HFF00FF
    m_Colors(27) = &HFFFF&
    m_Colors(28) = &HFFFF00
    m_Colors(29) = &H800080
    m_Colors(30) = &H80&
    m_Colors(31) = &H808000
    m_Colors(32) = &HFF0000
    m_Colors(33) = &HFFCC00
    m_Colors(34) = &HFFFFCC
    m_Colors(35) = &HCCFFCC
    m_Colors(36) = &H99FFFF
    m_Colors(37) = &HFFCC99
    m_Colors(38) = &HCC99FF
    m_Colors(39) = &HFF99CC
    m_Colors(40) = &H99CCFF
    m_Colors(41) = &HFF6633
    m_Colors(42) = &HCCCC33
    m_Colors(43) = &HCC99&
    m_Colors(44) = &HCCFF&
    m_Colors(45) = &H99FF&
    m_Colors(46) = &H66FF&
    m_Colors(47) = &H996666
    m_Colors(48) = &H969696
    m_Colors(49) = &H663300
    m_Colors(50) = &H669933
    m_Colors(51) = &H3300&
    m_Colors(52) = &H3333&
    m_Colors(53) = &H3399&
    m_Colors(54) = &H663399
    m_Colors(55) = &H993333
    m_Colors(56) = &H333333
    '   Set the mousepointer back to normal...
    Screen.MousePointer = vbDefault
End Sub

Private Sub Class_Terminate()
    '   Free up memory before we go...
    Screen.MousePointer = vbDefault
    Set oChart = Nothing
    Set oSheet = Nothing
    Set oExcel = Nothing
End Sub

Public Property Get DataLayout() As Orienation
    DataLayout = m_DataLayout
End Property

Public Property Let DataLayout(Layout As Orienation)
    m_DataLayout = Layout
End Property

Public Property Get HasChartTitle() As Boolean
    HasChartTitle = m_HasChartTitle
End Property

Public Property Let HasChartTitle(ChartTitle As Boolean)
     m_HasChartTitle = ChartTitle
End Property

Public Property Get HasPlotLegend() As Boolean
    HasPlotLegend = m_HasPlotLegend
End Property

Public Property Let HasPlotLegend(PlotLegend As Boolean)
     m_HasPlotLegend = PlotLegend
End Property

Public Property Get HasXAxisTitle() As Boolean
    HasXAxisTitle = m_HasXAxisTitle
End Property

Public Property Let HasXAxisTitle(HasTitle As Boolean)
    m_HasXAxisTitle = HasTitle
End Property

Public Property Get HasXMajorGrid() As Boolean
    HasXMajorGrid = m_HasXMajorGrid
End Property

Public Property Let HasXMajorGrid(HasGrid As Boolean)
    m_HasXMajorGrid = HasGrid
End Property

Public Property Get HasXMinorGrid() As Boolean
    HasXMinorGrid = m_HasXMinorGrid
End Property

Public Property Let HasXMinorGrid(HasGrid As Boolean)
    m_HasXMinorGrid = HasGrid
End Property

Public Property Get HasYAxisTitle() As Boolean
    HasYAxisTitle = m_HasYAxisTitle
End Property

Public Property Let HasYAxisTitle(HasTitle As Boolean)
    m_HasYAxisTitle = HasTitle
End Property

Public Property Get HasYMajorGrid() As Boolean
    HasYMajorGrid = m_HasYMajorGrid
End Property

Public Property Let HasYMajorGrid(HasGrid As Boolean)
    m_HasYMajorGrid = HasGrid
End Property

Public Property Get HasYMinorGrid() As Boolean
    HasYMinorGrid = m_HasYMinorGrid
End Property

Public Property Let HasYMinorGrid(HasGrid As Boolean)
    m_HasYMinorGrid = HasGrid
End Property

Public Property Get InteriorColorIndex() As Long
    '   Plot interior color
    InteriorColorIndex = m_InteriorColorIndex
End Property

Public Property Let InteriorColorIndex(Color As Long)
    '   Check to see if we are within range
    If (Color >= 1) And (Color < 57) Then
        m_InteriorColorIndex = Color
    Else
        '   We are not so tell the user
        Err.Raise 9, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Property

Public Property Get InteriorColorVal() As Long
    '   Plot interior color
    InteriorColorVal = m_InteriorColorVal
End Property

Public Property Let InteriorColorVal(Color As Long)
    '   Plot interior color
    m_InteriorColorVal = Color
End Property

Public Property Set OLEContainer(ByRef OLEObj As OLE)
    '   Map the container so we can use it later
    Set oOLE = OLEObj
End Property

Public Sub Plot(XArr() As Double, YArr() As Double, XErr() As Double, YErr() As Double)
    '   XDATA = XARR, YDATA = YARR, XERRORBAR_DATA = XErr, YERRORBAR_DATA = YErr
    Dim i                   As Integer      'Loop Counter
    Dim LBX1                As Integer      'Number of Current Rows for X in Dim 1
    Dim LBX2                As Integer      'Number of Current Rows for X in Dim 2
    Dim LBXE1               As Integer      'Number of Current Rows for X Errors in Dim 1
    Dim LBXE2               As Integer      'Number of Current Rows for X Errors in Dim 2
    Dim LBY1                As Integer      'Number of Current Rows for Y in Dim 1
    Dim LBY2                As Integer      'Number of Current Rows for Y in Dim 2
    Dim LBYE1               As Integer      'Number of Current Rows for Y Errors in Dim 1
    Dim LBYE2               As Integer      'Number of Current Rows for Y Errors in Dim 2
    Dim ScaleFactor         As Double       'Scale Factor to reduce the Chart Area
    Dim UBX1                As Integer      'Number of Current Cols for X in Dim 1
    Dim UBX2                As Integer      'Number of Current Cols for X in Dim 2
    Dim UBXE1               As Integer      'Number of Current Cols for X Errors in Dim 1
    Dim UBXE2               As Integer      'Number of Current Cols for X Errors in Dim 2
    Dim UBY1                As Integer      'Number of Current Cols for Y in Dim 1
    Dim UBY2                As Integer      'Number of Current Cols for Y in Dim 2
    Dim UBYE1               As Integer      'Number of Current Cols for Y Errors in Dim 1
    Dim UBYE2               As Integer      'Number of Current Cols for Y Errors in Dim 2
    Dim UnionRange          As Range        'Range Object for the Union Range of X & Y Data
    Dim XData()             As Double       'Final X Array to Pass to the Range Object
    Dim XError()            As Double       'Final X Error Bar Array to pass to the Range Object
    Dim XErrorRange         As Range        'Range Object for the XError Values
    Dim YData()             As Double       'Final Y Array to Pass to the Range Object
    Dim YError()            As Double       'Final Y Error Bar Array to pass to the Range Object
    Dim YErrorRange         As Range        'Range Object for the YError Values
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Turn off the screen updating to prevent flicker
    Call LockWindowUpdate(oOLE.Parent.hWnd)
    '   Embed a new Excel Chart into the OLE2 Control:
    oOLE.CreateEmbed "", "Excel.Chart"
    '   Check for versions....as Version 7 & 8 are not handled the same
    '   Set object references to Chart, Worksheet, and Application objects
    If Left(oOLE.object.Application.Version, 1) = "7" Then
        '   Excel 95's object model is different from Excel 97's
        Set oChart = oOLE.object    '   Chart1 Default Chart
    Else
        '   Assume all future Excel object models are going to be backward
        '   compatable with the Excel 97 the workbook
        Set oChart = oOLE.object.ActiveChart
    End If
    '   Sheet1 Default Data
    Set oSheet = oChart.Parent.Worksheets(1)
    
    ' Get the broundries of the data arrays passed
    LBX1 = LBound(XArr, 1)
    UBX1 = UBound(XArr, 1)
    LBX2 = LBound(XArr, 2)
    UBX2 = UBound(XArr, 2)
    LBY1 = LBound(YArr, 1)
    UBY1 = UBound(YArr, 1)
    LBY2 = LBound(YArr, 2)
    UBY2 = UBound(YArr, 2)
    ' Get the broundries of the error data arrays passed
    LBXE1 = LBound(XErr, 1)
    UBXE1 = UBound(XErr, 1)
    LBXE2 = LBound(XErr, 2)
    UBXE2 = UBound(XErr, 2)
    LBYE1 = LBound(YErr, 1)
    UBYE1 = UBound(YErr, 1)
    LBYE2 = LBound(YErr, 2)
    UBYE2 = UBound(YErr, 2)
    '   Clear the Contents of Worksheet
    oSheet.Cells.ClearContents
    
    '   Pass the data onto the Worksheet:
    With oSheet
        '   Make sure we do not omit the second dimension (i.e. as in vectors),
        '   or we will only get a single value passed to the range object!
        '   Make a local copy, since arrays are passed by reference we do not
        '   want any changes made to them while passing to the range objects...
        If (LBX2 = 0) And (UBX2 = 0) Then
            LBX2 = 1
            UBX2 = 1
            ReDim XData(LBX1 To UBX1, 1 To 1)
            For i = LBX1 To UBX1
                XData(i, 1) = XArr(i)
            Next i
        Else
            '   Must be 2D already, so use it as is...
            XData = XArr
        End If
        '   Now for XErr bar data
        If (LBXE2 = 0) And (UBXE2 = 0) Then
            LBXE2 = 1
            UBXE2 = 1
            ReDim XError(LBXE1 To UBXE1, 1 To 1)
            For i = LBXE1 To UBXE1
                XError(i, 1) = XErr(i)
            Next i
        Else
            '   Must be 2D already, so use it as is...
            XError = XErr
        End If
        '   YData range...
        If (LBY2 = 0) And (UBY2 = 0) Then
            LBY2 = 1
            UBY2 = 1
            ReDim YData(LBY1 To UBY1, 1 To 1)
            For i = LBY1 To UBY1
                YData(i, 1) = YArr(i)
            Next i
        Else
            '   Must be 2D already, so use it as is...
            YData = YArr
        End If
        '   YData range...
        If (LBYE2 = 0) And (UBYE2 = 0) Then
            LBYE2 = 1
            UBYE2 = 1
            ReDim YError(LBYE1 To UBYE1, 1 To 1)
            For i = LBYE1 To UBYE1
                YError(i, 1) = YErr(i)
            Next i
        Else
            '   Must be 2D already, so use it as is...
            YError = YErr
        End If
        '   Pass the data to the worksheet in the specified range
        .Range(.Cells(LBX1, LBX2), .Cells(UBX1, UBX2)).Value = XData
        .Range(.Cells(LBY1, LBX2 + LBY2), .Cells(UBY1, UBX2 + UBY2)).Value = YData
        '   Pass the XError data to the worksheet in the adjoining columns
        .Range(.Cells(LBY1, UBX2 + UBY2 + LBXE2), .Cells(UBY1, UBX2 + UBY2 + UBXE2)).Value = XError
        '   Now set the range so we can pass to the chart object..
        Set XErrorRange = .Range(.Cells(LBY1, UBX2 + UBY2 + LBXE2), .Cells(UBY1, UBX2 + UBY2 + UBXE2))
        '   Pass the YError data to the worksheet in the adjoining columns
        .Range(.Cells(LBY1, UBX2 + UBY2 + UBXE2 + LBYE2), .Cells(UBYE1, UBX2 + UBY2 + UBXE2 + UBYE2)).Value = YError
        '   Now set the range so we can pass to the chart object..
        Set YErrorRange = .Range(.Cells(LBY1, UBX2 + UBY2 + UBXE2 + LBYE2), .Cells(UBYE1, UBX2 + UBY2 + UBXE2 + UBYE2))
    End With
    '   Set the Union Range Object...This is how we pass the two ranges
    '   to the Chart Object as a single Range
    Set UnionRange = oSheet.Range(oSheet.Cells(LBX1, LBX2), oSheet.Cells(UBY1, UBX2 + UBY2))
    '   Since not all Chart Objects are supported in all plot types,
    '   we need to quietly bypass the exceptions...
    On Error Resume Next
    '   Set the boarder color of the chart by setting the OLE backcolor.
    '   Normally, one would set the color within the Excel Chart object, but
    '   since it is hosted by OLE this does not work the same.
    With oChart
        '   Setup the initial chart styles, and pass the data...
        '
        '   See if we are using color index or value for the borders
        If ChartColorVal = -1 Then
            '   Chart border color change by Index into Excel Color Table
            .ChartArea.Interior.ColorIndex = ChartColorIndex
            '   This prevent any uneven boarder from appearing in OLE2 container
            '   by setting the OLE2 container to the same color from our local
            '   replicate of the Excel Color table.
            oOLE.BackColor = m_Colors(ChartColorIndex)
        Else
            '   The following could be used instead of the index,
            '   as long as one passes Hex color values...also the OLE2 Container needs
            '   the same value to prevent a border effect around the chart.
            .ChartArea.Interior.Color = ChartColorVal
            '   This prevent any uneven boarder from appearing in OLE2 container
            oOLE.BackColor = ChartColorVal
        End If
        '   Do we have a title, if not then supress the display
        .HasTitle = HasChartTitle
        '   Pass the title text
        .ChartTitle.Characters.Text = ChartTitle
        '   Change the font to a usable value
        .ChartTitle.Font.Size = ChartTitleSize
        '   Set the chart style
        .ChartType = ChartStyle
        .HasLegend = HasPlotLegend
        '   Pass the plot union-range data to the chart ojbect
        .SetSourceData Source:=UnionRange, PlotBy:=xlColumns
        '   See if we want X-Error Bars on the plot
        If (XErrorRange Is Nothing) = False Then
            If (XErrorRange.Columns.Count = 1) And (YErrorRange.Columns.Count) > 1 Then
                '   Handle cases when we only have 1 X Range and Mulitple Y
                For i = 1 To YErrorRange.Columns.Count
                    '   Select the active chart
                    oChart.SeriesCollection(i).Select
                    '   Add X Error Bars
                    oChart.SeriesCollection(i).ErrorBar Direction:=xlX, Include:=xlBoth, _
                        Type:=xlCustom, Amount:=XErrorRange.Columns(1), MinusValues:=XErrorRange.Columns(1)
                Next i
            Else
                For i = 1 To XErrorRange.Columns.Count
                    '   Select the active chart
                    oChart.SeriesCollection(i).Select
                    '   Add X Error Bars
                    oChart.SeriesCollection(i).ErrorBar Direction:=xlX, Include:=xlBoth, _
                        Type:=xlCustom, Amount:=XErrorRange.Columns(i), MinusValues:=XErrorRange.Columns(i)
                Next i
            End If
            
        End If
        '   See if we want Y-Error Bars on the plot
        If (YErrorRange Is Nothing) = False Then
            For i = 1 To YErrorRange.Columns.Count
                '   Select the active chart
                oChart.SeriesCollection(i).Select
                '   Add Y Error Bars
                oChart.SeriesCollection(i).ErrorBar Direction:=xlY, Include:=xlBoth, _
                    Type:=xlCustom, Amount:=YErrorRange.Columns(i), MinusValues:=YErrorRange.Columns(i)
            Next i
        End If
        '   Set the X Axis Title, Gridlines, Scaling and Plot Options
        With .Axes(xlCategory, xlPrimary)
            '   Supress the X Axis title?
            .HasTitle = HasXAxisTitle
            '   Pass the value to the object
            .AxisTitle.Characters.Text = XAxisTitle
            '   Set the size of the font
            .AxisTitle.Font.Size = XLabelSize
            '   Let the chart figure out where this should cross
            .Crosses = xlAutomatic
            .DisplayUnit = xlNone
            '   Gridlines (major, minor)
            .HasMajorGridlines = HasXMajorGrid
            .HasMinorGridlines = HasXMinorGrid
            .MajorUnitIsAuto = XAxisMajorUnitsAuto
            .MaximumScaleIsAuto = XAxisMaxScaleAuto
            .MinimumScaleIsAuto = XAxisMinScaleAuto
            .MinorUnitIsAuto = XAxisMinorUnitsAuto
            '   Set the data up in reverse order of passed array
            .ReversePlotOrder = ReversePlotOrder
            '   Set the X Scale type (linear, log10)
            .ScaleType = XAxisScaleType
            '   Set the size of the tick labels
            .TickLabels.Font.Size = XTickLabelSize
        End With
        '   Set the Y Axis Title, Gridlines, Scaling and Plot Options
        With .Axes(xlValue, xlPrimary)
            '   Supress the Y Axis title?
            .HasTitle = HasYAxisTitle
            '   Pass the value to the object
            .AxisTitle.Characters.Text = YAxisTitle
            '   Set the size of the font
            .AxisTitle.Font.Size = YLabelSize
            '   Let the chart figure out where this should cross
            .Crosses = xlAutomatic
            .DisplayUnit = xlNone
            '   Gridlines (major, minor)
            .HasMajorGridlines = HasYMajorGrid
            .HasMinorGridlines = HasYMinorGrid
            .MajorUnitIsAuto = YAxisMajorUnitsAuto
            .MaximumScaleIsAuto = YAxisMaxScaleAuto
            .MinimumScaleIsAuto = YAxisMinScaleAuto
            .MinorUnitIsAuto = YAxisMinorUnitsAuto
            '   Set the data up in reverse order of passed array
            .ReversePlotOrder = ReversePlotOrder
            '   Set the Y Scale type (linear, log10)
            .ScaleType = YAxisScaleType
            '   Set the size of the tick labels
            .TickLabels.Font.Size = XTickLabelSize
        End With
        '   Set Plot Area Color and Pattern index
        With .PlotArea
            '   Clear the current plot formats to start
            .ClearFormats
            '   Compute the Scale Factor
            If Val(oOLE.object.Application.Version) <= "9" Then
                '   In Excel 9 or less the size of the plot is not the
                '   same, so we need to adjust the Chart size
                ScaleFactor = 0.65
            Else
                '   Using Excel 10 (XP) or greater so adjsut accordingly...
                ScaleFactor = 0.75
            End If
            '   Make sure to fill the OLE field as much as
            '   possible by scaling this to the container dimensions
            If Screen.ActiveForm.ScaleMode = 1 Then
                '   The form is using Twips, so convert to pixels
                .Height = Screen.ActiveForm.ScaleY(oOLE.Height, Screen.ActiveForm.ScaleMode, vbPixels) * (ScaleFactor - 0.05)
                .Width = Screen.ActiveForm.ScaleX(oOLE.Width, Screen.ActiveForm.ScaleMode, vbPixels) * ScaleFactor
            Else
                '   The form is using Pixels, so use as is...
                .Height = oOLE.Height * ScaleFactor
                .Width = oOLE.Width * ScaleFactor
            End If
            .Top = oOLE.Top * 0.12
            '   Set the interior of the plot
            With .Interior
                '   Set the color of the Plot Area....
                '   See if we have set the Long value of the color,
                '   or is it an Index into Excel's Color table.
                If m_InteriorColorVal = -1 Then
                    '   We are using the Excel Color Table...
                    .ColorIndex = InteriorColorIndex
                Else
                    '   The following could be used instead of the index,
                    '   as long as one passes Long color values
                    .Color = InteriorColorVal
                End If
                .PatternColorIndex = 1
                '   Use solid backgrounds which appear to be
                '   the cleaner than textures...
                .Pattern = 1
            End With
            '   Eliminate the Boarder as this makes the Chart and OLE2
            '   container appear as two object instead of as one...
            With .Border
                .LineStyle = -1
                .Weight = 1
            End With
        End With
    End With
    '   Turn on error handling again....
    On Error GoTo 0
    '   Turn On the screen updating....
    Call LockWindowUpdate(0)
    
Sub_ErrHandlerExit:
    '   Free up memory before we go...
    Set oChart = Nothing
    Set oSheet = Nothing
    Exit Sub
Sub_ErrHandler:
    '   See if the array index exceeds boundries
    If Err.Number = 9 Then
        Resume Next
    Else
        Err.Raise Err.Number, "CExcelPlot.Plot", Err.Description, Err.HelpFile, Err.HelpContext
        Resume Sub_ErrHandlerExit:
    End If
End Sub

Public Property Get ReversePlotOrder() As Boolean
    ReversePlotOrder = m_ReversePlotOrder
End Property

Public Property Let ReversePlotOrder(ReverseOrder As Boolean)
     m_ReversePlotOrder = ReverseOrder
End Property

Public Property Get XAxisMajorUnitsAuto() As Boolean
    XAxisMajorUnitsAuto = m_XAxisMajorUnitsAuto
End Property

Public Property Let XAxisMajorUnitsAuto(UnitsAuto As Boolean)
     m_XAxisMajorUnitsAuto = UnitsAuto
End Property

Public Property Get XAxisMaxScaleAuto() As Boolean
    XAxisMaxScaleAuto = m_XAxisMaxScaleAuto
End Property

Public Property Let XAxisMaxScaleAuto(ScaleAuto As Boolean)
     m_XAxisMaxScaleAuto = ScaleAuto
End Property

Public Property Get XAxisMinorUnitsAuto() As Boolean
    XAxisMinorUnitsAuto = m_XAxisMinorUnitsAuto
End Property

Public Property Let XAxisMinorUnitsAuto(UnitsAuto As Boolean)
     m_XAxisMinorUnitsAuto = UnitsAuto
End Property

Public Property Get XAxisMinScaleAuto() As Boolean
    XAxisMinScaleAuto = m_XAxisMinScaleAuto
End Property

Public Property Let XAxisMinScaleAuto(ScaleAuto As Boolean)
     m_XAxisMinScaleAuto = ScaleAuto
End Property

Public Property Get XAxisScaleType() As ScaleMode
    XAxisScaleType = m_XAxisScaleType
End Property

Public Property Let XAxisScaleType(ScaleType As ScaleMode)
     m_XAxisScaleType = ScaleType
End Property

Public Property Get XAxisTitle() As String
    XAxisTitle = m_XAxisTitle
End Property

Public Property Let XAxisTitle(AxisTitle As String)
    m_XAxisTitle = AxisTitle
End Property

Public Property Get XLabelSize() As String
    XLabelSize = m_XLabelSize
End Property

Public Property Let XLabelSize(LabelSize As String)
    m_XLabelSize = LabelSize
End Property
    
Public Property Get XTickLabelSize() As String
    XTickLabelSize = m_XTickLabelSize
End Property

Public Property Let XTickLabelSize(LabelSize As String)
    m_XTickLabelSize = LabelSize
End Property

Public Property Get YAxisMajorUnitsAuto() As Boolean
    YAxisMajorUnitsAuto = m_YAxisMajorUnitsAuto
End Property

Public Property Let YAxisMajorUnitsAuto(UnitsAuto As Boolean)
     m_YAxisMajorUnitsAuto = UnitsAuto
End Property

Public Property Get YAxisMaxScaleAuto() As Boolean
    YAxisMaxScaleAuto = m_YAxisMaxScaleAuto
End Property

Public Property Let YAxisMaxScaleAuto(ScaleAuto As Boolean)
     m_YAxisMaxScaleAuto = ScaleAuto
End Property

Public Property Get YAxisMinorUnitsAuto() As Boolean
    YAxisMinorUnitsAuto = m_YAxisMinorUnitsAuto
End Property

Public Property Let YAxisMinorUnitsAuto(UnitsAuto As Boolean)
     m_YAxisMinorUnitsAuto = UnitsAuto
End Property

Public Property Get YAxisMinScaleAuto() As Boolean
    YAxisMinScaleAuto = m_YAxisMinScaleAuto
End Property

Public Property Let YAxisMinScaleAuto(ScaleAuto As Boolean)
     m_YAxisMinScaleAuto = ScaleAuto
End Property

Public Property Get YAxisScaleType() As ScaleMode
    YAxisScaleType = m_YAxisScaleType
End Property

Public Property Let YAxisScaleType(ScaleType As ScaleMode)
     m_YAxisScaleType = ScaleType
End Property

Public Property Get YAxisTitle() As String
    YAxisTitle = m_YAxisTitle
End Property

Public Property Let YAxisTitle(AxisTitle As String)
    m_YAxisTitle = AxisTitle
End Property

Public Property Get YLabelSize() As String
    YLabelSize = m_YLabelSize
End Property

Public Property Let YLabelSize(LabelSize As String)
    m_YLabelSize = LabelSize
End Property
    
Public Property Get YTickLabelSize() As String
    YTickLabelSize = m_YTickLabelSize
End Property

Public Property Let YTickLabelSize(LabelSize As String)
    m_YTickLabelSize = LabelSize
End Property



