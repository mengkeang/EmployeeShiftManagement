Option Explicit

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' We need to work on these two charts as we put the first serie as the target value
    '   on Secondary Axis. Otherwise, the visual correctness is not maintained.
    
    If Target.Name = "pvEmployeeWorkingHourVsContractHour" Then
        SyncAxisesForChart ChartName:="chtWorkingHourVsContractHour"
    End If
    
    If Target.Name = "pvPremiumShiftDistribution" Then
        SyncAxisesForChart ChartName:="chtPremiumShift"
    End If
    
    ' Enable screen after we finish the code above
    Application.ScreenUpdating = True
End Sub

Sub SyncAxisesForChart(ChartName As String)
    Dim chtObj As ChartObject
    Dim cht As Chart
    
    On Error Resume Next
    Set chtObj = SheetDashboard.ChartObjects(ChartName)
    
    On Error GoTo FINALLY
    If Not chtObj Is Nothing Then
        Set cht = chtObj.Chart
        
        ' Set the first series as Secondary Axis
        ' We setup the first Series as the Target Value
        cht.FullSeriesCollection(1).AxisGroup = 2
        
        ' Make sure the Secondary Axis Min and Max Scale = Primary Axis Min and Max Scale
        ' This will make the Target Value of First Series on the same scale as other series
        With cht.Axes(xlValue, xlSecondary)
            .MaximumScale = cht.Axes(xlValue, xlPrimary).MaximumScale
            .MinimumScale = cht.Axes(xlValue, xlPrimary).MinimumScale
            .TickLabelPosition = xlTickLabelPositionNone ' Keeps it clean/hidden
        End With
    End If
    
FINALLY:
    Set cht = Nothing
    Set chtObj = Nothing
End Sub
