# Introduction
Meet Bob. He owns a busy restaurant. The restaurant opens from 8AM until 11PM every day including weekends and holidays.He has manager, assistant manager, chefs, assistant chef, and server. Manager and assistant manager are keyholders to open and close the restaurant. Full time staff works up to 40 hours per week, while part time staff work up to 20 hours per week. For weekends and holidays, the pay rate is 1.5 and 2.0 respectively. These are premium shifts.

Everything seems fine. But behind the scenes, it’s chaos. Every week, he creates staff schedules manually. Some staff work too many hours. Others not enough. There’s no system to track total weekly working hours for full time and part time staff.

Then comes payroll. Weekend and holiday rates are different and everything is calculated manually.
Sometimes, a staff works late at night and comes back early the next morning. That’s called **CLOPEN**.
It’s exhausting. It’s unfair. Bob doesn’t even realize it until it happens.

Staff start complaining why certain names are scheduled on most of the premium shifts. Bob wants to be fair. However, without data, he’s just guessing.

Bob also forgets to assign manager to open the restaurant in some instance…
He starts to feel the burden of all of these.

# Features to be developed
we build the employee shirt management system using Excel covering the following features:
* Weekly wage calculation for each staff
* Weekly contract hour detection
* Premium Shift summary of individual staff
* CLOPEN detection
* Daily staffing headcount on critical areas of the restaurant
* An interactive dashboard for Bob
* Staff Information Management
![staff list](img/staff_list.png)

* Public Holiday List Customization
![holiday list](img/holiday_list.png)

* Archive Weekly Data for Dashboard and Analysis
![archive weekly data](img/archive_weekly_data.png)

* Configuration on Weekend and Holiday Wage Rate
![configuration](img/configuration.png)

> *The solution will work on older version of Excel. We do not use spilled array and modern formula from newer version of Excel.*

# tblStaff vs tblSchedule
We have two important tables in our system: `tblStaff` and `tblSchedule`. `tblStaff` is on Sheet Staff, while `tblSchedule` is on Sheet Schedule (our main UI for Bob to manage the employee shift).

`tblSchedule` retrieve all information from tblStaff based on the drop-down list on column `tblSchedule[StaffName]`.

![tblStaff vs tblSchedule](img/tblStaff_vs_tblSchedule.png)

There are two reasons on such decision to separate the same information into two tables.

1. I want `tblSchedule` to have minimal information about staff. The main information is about shift assignment and management. Putting all staff information can make the main UI messy and hard for Bob to operate.
2. `tblStaff` can contain more staff than `tblSchedule`. This can happen when a staff resigned from the role and leave the restaurant. `tblSchedule` should include only active staff. Separation data makes more sense for Bob.

## New Staff Button
To make things easy for Bob, we introduce a button `New Staff` to make sure the list of staff sync with `tblSchedule`.

Here is the VBA code for the button `New Staff`.

```vb
Sub btnNewStaff_Click()
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Add New Row to tblStaff
    ' --------------------------------------------------------------------------------
    Dim tblStaff As ListObject
    Set tblStaff = SheetStaff.ListObjects("tblStaff")
    
    ' Add a new row on tblStaff
    Dim newRow As ListRow
    Set newRow = tblStaff.ListRows.Add
    
    ' Select Column: StaffName of the newly inserted row
    newRow.Range(tblStaff.ListColumns("StaffName").Index).Select
    
    ' --------------------------------------------------------------------------------
    ' Step: Add New Row to tblSchedule and Insert an empty row below the newly added row
    ' --------------------------------------------------------------------------------
    Dim tblSchedule As ListObject
    Set tblSchedule = SheetSchedule.ListObjects("tblSchedule")
    
    ' Add a new row on tblSchedule
    Dim newRowSchedule As ListRow
    Set newRowSchedule = tblSchedule.ListRows.Add
    
    ' Find the Row Number of the newly inserted row
    '   Plus 2 so that there are gaps between tblSchedule and staff head count summary section
    Dim nextRow As Long
    nextRow = newRowSchedule.Range().Row + 2
    
    ' Insert a row at index: nextRow
    SheetSchedule.Range("A" & nextRow).EntireRow.Insert
    
    ' Enable Screen Update after we complete our task
    Application.ScreenUpdating = True
End Sub
```

# tblRefShift
`tblRefShift` is on Sheet Ref. It is mainly used on `tblSchedule`. The column `ShiftCode` is used as the drop-down list on `tblSchedule`.

Our main goal is that when Bob selects the `ShiftCode` on `tblSchedule`, we can use formula to find `Hour` for the assigned shift.

![tblRefShift](img/tblRefShift.png)

# Sheet Schedule: Archive Button
We prepare a button to achieve weekly schedule data into a history data sheet.

We need history data for two main reasons:
1. We need history data to calculate the Premium Shift Count per employee. This will allow to detect which staff was assigned the premium shift more frequently than the other.
2. Historical data are needed for our interactive dashboard.

To prepare the historical data, we need to perform data transformation using Power Query and VBA code.

## Power Query: tblSchedule to tblWeeklyData
You may notice that `tblSchedule` is structured as wide table (many horizontal columns). It is easy for Bob to operate and assign the shift for his staff. Such structure is not dashboard friendly. We will need to transform tblSchedule into long table (more vertical rows).

To convert from wide table to long table, we need to do unpivot operation on tblSchedule. Since we design the schedule as Excel table already, we can perform the unpivot operation in Power Query easily.

> Unpivot Column: turn selected columns into rows. After the operation, there will be two columns: Attribute and Value. Column Attribute will contain the column label. The Value column will be value of the unpivot columns.

![PowerQuery tblWeeklyData](img/PowerQuery_tblWeeklyData.png)

![unpivot tblWeeklyData](img/unpivot_tblWeeklyData.png)

## tblWeeklyData: Calculated Columns
After we complete the PowerQuery, we will need to add more information for our dashboard.

We add the following columns:
* **Date**: `=INDEX(Schedule!$9:$9, MATCH([@WeekDay], tblSchedule[#Headers],0))`
* **IsWeekend**: `=IF(WEEKDAY([@Date], 2)>=6, "YES", "NO")`
* **IsHoliday**: `=IF(NOT(ISERROR(MATCH([@Date], tblPublicHoliday[HOLIDAY], 0))), "YES", "NO")`
* **HolidayOrWeekend**: `=IF(OR([@IsWeekend]="YES", [@IsHoliday]="YES"), "YES", "NO")`
* **WorkingHour**: `=INDEX(tblRefShift[Hour], MATCH(CONCAT([@Role],"_", [@ShiftCode]), tblRefShift[ShiftID], 0))`
* **Wage**: `=IF([@IsHoliday]="YES", holiday_rate_multiplier, IF(WEEKDAY([@Date], 2)=6, sat_rate_multiplier, IF(WEEKDAY([@Date], 2)=7, sun_rate_multiplier, 1))) * [@HourlyRate]*[@WorkingHour]`
* **MaxHourPerDay**: `=[@MaxHour]/7`
* **PremiumShift**: `=IF(AND([@HolidayOrWeekend]="YES", [@WorkingHour]>0), 1, 0)`

You may notice column: `MaxHourPerDay`. Since `tblWeeklyData` is on daily unit, we need to convert the `MaxHour` per week into daily unit by dividing by 7.

Column: `PremiumShift` is the one we need to count the PremiumShift per employee.

![tblWeekly more columns](img/tblWeekly_more_columns.png)

> Note: Power Query can be used to get all the calculated columns above. However, I think using Excel formula is more elegant in a sense that it works fast with such small dataset.

## Sheet tblHistory: tblHistory
After we complete the data transformation on `tblWeeklyData`, we prepare another sheet for historical data.

We prepared `tblHistory` by copying everything from `tblWeeklyData`. We renamed the table to `tblHistory`.

Our main purpose is to copy data from `tblWeeklyData` and paste them on `tblHistory`. We will paste value only. Therefore, `tblHistory` does not contain any formula.

## Archive Button
With everything is in place, our final step is to write a VBA procedure to archive the data from tblWeeklyData to tblHistory.

```vb
Sub btnArchiveData_Click()
    
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Copy Column: Sun to Column: LW_Sun
    ' --------------------------------------------------------------------------------
    ' This is for Clopen Detection Conditional Formatting
    Dim tblSchedule As ListObject
    Set tblSchedule = SheetSchedule.ListObjects("tblSchedule")
    
    ' Copy Data from Column: Sun to Column: LW_Sun
    tblSchedule.ListColumns("Sun").DataBodyRange.Copy
    tblSchedule.ListColumns("LW_Sun").DataBodyRange.PasteSpecial xlPasteValues
    
    ' End Copy and Paste
    Application.CutCopyMode = False
    
    ' --------------------------------------------------------------------------------
    ' Step: Refresh tblWeeklyData
    ' --------------------------------------------------------------------------------
    Dim tblWeeklyData As ListObject
    Set tblWeeklyData = SheettblWeeklyData.ListObjects("tblWeeklyData")
    
    ' Refresh data on Sheet tblWeeklyData via Power Query
    tblWeeklyData.Refresh

    ' --------------------------------------------------------------------------------
    ' Step: Copy Data from tblWeeklyData and Append to tblHistory
    ' --------------------------------------------------------------------------------
    Dim tblHistory As ListObject
    Set tblHistory = SheettblHistory.ListObjects("tblHistory")
    
    ' Add new row to tblHistory
    ' This is the row where will paste the data from tblWeeklyData
    Dim newRow As ListRow
    Set newRow = tblHistory.ListRows.Add
    
    ' Copy data from Sheet tblWeeklyData
    ' Paste data on Sheet tblHistory > New Row
    tblWeeklyData.DataBodyRange.Copy
    newRow.Range(tblHistory.ListColumns("StaffName").Index).PasteSpecial xlPasteValues
    
    ' End Copy and Paste
    Application.CutCopyMode = False
    ' --------------------------------------------------------------------------------
    
    ' Show successful message
    MsgBox "Action Completed!", vbInformation, "AbcFunWithExcel"
    
    ' Enable Screen Update after we complete our task
    Application.ScreenUpdating = True
    
End Sub
```

# Premium Shift Count
Now we have tblHistory in hand. Then we can find the premium shift count of each employee based on the historical data.

Because we design the solution to work with older version of Excel, things are more challenging to find the premium shift count.

For this reason, we will need to work out by creating new measure using DAX. DAX is part of the Power Pivot add-in with Excel.

## Create a Power Pivot Table

We created a Power Pivot Table on Sheet `Setting` at Range: `G17`. The data soure is from `tblHistory`. 

Make sure to tick **"Add this data to the Data Model"**. This will create a Power Pivot Table instead of a normal Pivot Table.

![create power pivot premiumshift](img/create_power_pivot_premiumshift.png)

Then drag field: `StaffName` and `Date` to Rows section, PremiumShift to Value section.

![power_pivot_premiumshift](img/power_pivot_premiumshift.png)

We want to reset the Premium Shift Count every year. To do this, we can apply filter on Date Column and set the filter to `This Year`. Later, in case Bob would like to reset the Premium Shift Count by Quarter or Semester, we can change the filter here easily.

## Add New Measure: HolidayWeekendCount
Then we need to add a new measure by right click on `tblHistory` and select "Add Measure..."

![power pivot addmeasure](img/power_pivot_addmeasure.png)

Measure name: HolidayWeekendCount
DAX: `=CALCULATE(DISTINCTCOUNT([Date]), KEEPFILTERS(tblHistory[HolidayOrWeekend]="YES"))`

![dax holidayweekendcount](img/dax_holidayweekendcount.png)

Now we get the Premium Shift Count per employee along with the total Premium Shift within the `tblHistory` dataset. We created a named range by pointing to the total Premium Shift Count. We named it as `holiday_weekend_count`.

![power pivot TotalPremiumShift](./img/power_pivot_TotalPremiumShift.png)

## Display Premium Shift on tblSchedule
Final step is to display the Premium Shift Count per staff and total Premium Shift on `tblSchedule`.

We use this formula:

```
=CONCAT(IFERROR(GETPIVOTDATA("[Measures].[Sum of PremiumShift]",Setting!$G$17,"[tblHistory].[StaffName]","[tblHistory].[StaffName].&[" & [@StaffName] & "]"), 0), " / ", holiday_weekend_count)
```

![tblSchedule PremiumShift](./img/tblSchedule_PremiumShift.png)

# Interactive Dashboard
Our final goal is to make an interactive dashboard. We prepared the history data for this purpose. We are not going to explain every step of how we design the whole dashboard. However, we will highlight key points about our decision and the chart design choice.

## Sheet Pivot and Sheet Dashboard
Our interactive dashboard is fully designed with Power Pivot using the tblHistory as the data source. We put all Power Pivot tables on Sheet Pivot.

Then we link the Power Pivot Tables to create various charts for Sheet Dashboard. We also add slicers to make our dashboard interactive to user input.

![Sheet Pivot](./img/SheetPivot.png)
![Sheet Dashboard](./img/SheetDashboard.png)

## Working Hours vs Contracted Hours
One of the important information is to know which staff had worked over their allowed working hours. We can use the following bar chart to show the target vs the actual value.

The chart below is good enough. However, it is not easy to understand at first glance. We can improve this chart with candle stick chart.

![WorkingHour vs ContractedHour before](./img/WorkingHour_vs_ContractedHour_before.png)

### Candle Stick Chart
We improve the above chart by introducing two measures: `SafeHour` and `OvertimedHour`. Then we add the two measures to the Power Pivot Table as shown in the image below. 

![WorkingHour vs ContractedHour after](./img/WorkingHour_vs_ContractedHour_after.png)

To produce the above chart, we use horizontal stack bar chart. We set the MaxHour (our target: allowed working hours) as the secondary axis.

![WorkingHour vs ContractedHour SecondaryAxis](./img/WorkingHour_vs_ContractedHour_SecondaryAxis.png)

![WorkingHour vs ContractedHour FormatMaxHour](./img/WorkingHour_vs_ContractedHour_FormatMaxHour.png)

![WorkingHour vs ContractedHour DeleteAxis](./img/WorkingHour_vs_ContractedHour_DeleteAxis.png)

We add Data Label to `OvertimedHour` data series. This will allow Bob to know exactly the number of overtimed hour of his staff. There are zero text display as well for staff who does not work more than the max hour. We can hide zero text using Custom Format on the data series.

![WorkingHour vs ContractedHour AddDataLabel](./img/WorkingHour_vs_ContractedHour_AddDataLabel.png)

We add the Custom Format `0;;;` to hide the zero text on the Data Label.
![WorkingHour vs ContractedHour CustomFormat](./img/WorkingHour_vs_ContractedHour_CustomFormat.png)

With these tweaks, our bar chart is more readable at the first glance.

![WorkingHour vs ContractedHour before after](./img/WorkingHour_vs_ContractedHour_before_after.png)

### Problems with Our Custom Chart
With interactive dashboard, we face issue on our custom chart. For example, there is data in Jun 2026 on tblHistory. When user clicks on the timeline to Jun 2026, there chart will be blank.

However, soon after the user clears filter on the timeline, the chart will be reverted back to its original format. That means MaxHour data series is on primary axis (not the secondary axis we set previously).

![chart problem](img/chart_problem.gif)

![CandleStickChart Problem](./img/CandleStickChart_Problem.png)

To solve this issue, we will need to write VBA code targeting the custom chart to ensure that MaxHour data series is always on the secondary axis.

### VBA Code on Pivot Table Update
We will need to give the chart a name and the Pivot Table a name as well. We will refer to the chart and pivot table object in VBA code later.

* Chart Name: `chtWorkingHourVsContractHour`
* Pivot Table Name: `pvEmployeeWorkingHourVsContractHour`

![Chart WorkingHourvsContractHour SetName](./img/Chart_WorkingHourvsContractHour_SetName.png)

![PivotTable pvEmployeeWorkingHour SetName](./img/PivotTable_pvEmployeeWorkingHour_SetName.png)

Then we will need to add the following code on Worksheet Event. Our target sheet is Sheet Pivot on event: PivotTableUpdate.

![SheetPivot PivotTableUpdate](./img/SheetPivot_PivotTableUpdate.png)

We used the following code:

```vb
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' We need to work on these two charts as we put the first serie as the target value
    '   on Secondary Axis. Otherwise, the visual correctness is not maintained.
    
    If Target.Name = "pvEmployeeWorkingHourVsContractHour" Then
        SyncAxisesForChart ChartName:="chtWorkingHourVsContractHour"
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

```

## Premium Shift Chart
Premium Shift Chart applied the same technique as Working Hours vs Contract Hours chart. We need three data series as following:
* Total Premium Shift Count (target value)
* Moderate (premium shift count per employee less than or equal to 75%)
* Frequent (premium shift count per employee greater than 75%)

![Premium Shift Count Chart](./img/PremiumShiftCountChart.png)

To achieve the above chart, we need to write DAX measures to calculate the Moderate and Frequent data series. 

[Total Premium Shift Count](#add-new-measure-holidayweekendcount): we already had the measure when we completed the Premium Shift Count feature.

### Moderate Data Series
Measure Name: `PremShiftBelow75`
DAX: `=IF(SUM([PremiumShift]) / [HolidayWeekendCount] <= 0.75, SUM([PremiumShift]), 0)`

![Measure PremiumShiftBelow75](./img/Measure_PremiumShiftBelow75.png)

> We decided to use 75% as the threshold percentage to differentiate between Moderate and Frequent. You can decide a different threshold value in the DAX formula.

### Frequent Data Series
Measure Name: `PremShiftOver75`
DAX: `=IF(SUM([PremiumShift]) / [HolidayWeekendCount] >0.75, SUM([PremiumShift]), 0)`

![Measure PremiumShiftOver75](./img/Measure_PremiumShiftOver75.png)

> We decided to use 75% as the threshold percentage to differentiate between Moderate and Frequent. You can decide a different threshold value in the DAX formula.

### Pivot Table: pvPremiumShiftDistribution

This is the pivot table for our Premium Shift Chart above. We will include `RoleCategory` and `StaffName` on the Rows section. Add measures to the pivot table.

![PivotTable pvPremiumShift](./img/PivotTable_pvPremiumShift.png)

Let's give the Pivot Table a name: `pvPremiumShiftDistribution`

### Create the Chart: chtPremiumShift
We will use the same technique to create this custom chart. You can follow the steps from this [section](#candle-stick-chart).

The main step for creating this custom chart is to set the first data series (Total Premium Shift) to secondary axis.

This will allow Total Premium Shift to overlap on other data series (Moderate and Frequent).

![Chart PremiumShift SecondaryAxis](./img/Chart_PremiumShift_SecondaryAxis.png)

We named this chart as `chtPremiumShift`.

### Problem with Our Custom Chart
`chtPremiumShift` chart shares the same problem as the one in [Working Hours vs Contracted Hours](#problems-with-our-custom-chart).

When user put filter(s) on the dashboard and there is no data for the applied filter, the chart will be empty. Once the user clears the filter(s), the secondary axis will revert to primary axis as we show in this image below.

![Chart PremiumShift](./img/Chart_PremiumShift.png)

To solve this issue, we will need to write the VBA code to target the Pivot Table: `pvPremiumShiftDistribution` and Chart: `chtPremiumShift`.

### VBA Code on Pivot Table Update
We already implemented the code for Working Hours vs Contracted Hours from the previous section. All we need to do for Premium Shift Chart is to add an IF block under `Worksheet_PivotTableUpdate`.

Here is the code snippet for `Worksheet_PivotTableUpdate`.

```vb
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    ' Disable Screen Update to boost performance
    Application.ScreenUpdating = False
    
    ' We need to work on these two charts as we put the first serie as the target value
    '   on Secondary Axis. Otherwise, the visual correctness is not maintained.
    
    If Target.Name = "pvEmployeeWorkingHourVsContractHour" Then
        SyncAxisesForChart ChartName:="chtWorkingHourVsContractHour"
    End If
    
    ' ADD These Section for chtPremiumShift
    If Target.Name = "pvPremiumShiftDistribution" Then
        SyncAxisesForChart ChartName:="chtPremiumShift"
    End If
    
    ' Enable screen after we finish the code above
    Application.ScreenUpdating = True
End Sub
```