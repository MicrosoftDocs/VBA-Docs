---
title: TimeScaleValue object (Project)
ms.prod: project-server
api_name:
- Project.TimeScaleValue
ms.assetid: bea0ad82-a3de-30d8-f191-dc2248c32653
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeScaleValue object (Project)

Represents a timescaled data item. The  **TimeScaleValue** object is a member of the **[TimeScaleValues](Project.timescalevalues.md)** collection.


## Examples

 **Using the TimeScaleValue Object**

Use  **TimeScaleValues** (_index_), where _index_ is the index number of the timescaled data item, to return a single **TimeScaleValue** object. The following example displays the number of hours of work per day for a resource during the first full week in October 2012.




```vb
Dim TSV As TimeScaleValues, HowMany As Long
Dim HoursPerDay As String

Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)

For HowMany = 1 To TSV.Count
    HoursPerDay = HoursPerDay & TSV(HowMany).StartDate & " - " & _
        TSV(HowMany).EndDate & ", " & TSV(HowMany) / 60 & vbCrLf
Next HowMany

MsgBox HoursPerDay
```

 **Using the TimeScaleValues Collection**

Use the  **[TimeScaleData](./Project.Resource.TimeScaleData.md)** method to return a **TimeScaleValues** collection. The following example returns a **TimeScaleValues** collection for the amount of work done by the resource in the active cell between the specified dates, split into week-long portions.




```vb
ActiveCell.Resource.TimeScaleData("10/1/2012", "10/31/2012")
```

Use the  **[Add](./Project.TimeScaleValues.Add.md)** method to add a **TimeScaleValue** object to the **TimeScaleValues** collection. The following example adds 8 hours of work to Tuesday of that week.




```vb
Dim TSV As TimeScaleValues

Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)
TSV.Add 480, 2
```


## Methods



|Name|
|:-----|
|[Clear](./Project.TimeScaleValue.Clear.md)|
|[Delete](./Project.TimeScaleValue.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.TimeScaleValue.Application.md)|
|[EndDate](./Project.TimeScaleValue.EndDate.md)|
|[Index](./Project.TimeScaleValue.Index.md)|
|[Parent](./Project.TimeScaleValue.Parent.md)|
|[StartDate](./Project.TimeScaleValue.StartDate.md)|
|[Value](./Project.TimeScaleValue.Value.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]