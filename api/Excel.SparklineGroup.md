---
title: SparklineGroup object (Excel)
keywords: vbaxl10.chm870072
f1_keywords:
- vbaxl10.chm870072
ms.prod: excel
api_name:
- Excel.SparklineGroup
ms.assetid: cc694d97-a3d3-3473-2e37-0ede67b97680
ms.date: 06/08/2017
localization_priority: Normal
---


# SparklineGroup object (Excel)

Represents a group of sparklines.

## Remarks

The  **SparklineGroup** object can contain multiple sparklines and contains the property settings for the group, such as color and axis settings. Each sparkline is represented by a **[Sparkline](Excel.Sparkline.md)** object.

Use the **[Modify](Excel.SparklineGroup.Modify.md)** method to add or remove sparklines from the sparkline group. Use the **[ModifyLocation](Excel.SparklineGroup.ModifyLocation.md)** method to change the location of the sparkline, and use the **[ModifySourceData](Excel.SparklineGroup.ModifySourceData.md)** method to change the range of the source data.

**Note**: Application.ReferenceStyle must be set to xlA1 to execute SparklineGroups.Add.

## Example

The following code example creates a group of column sparklines at the location A1:A4 that are bound to the source data in the range Sheet2!B1:E4. The series color is changed to display the columns in red.

```vb
Dim mySG As SparklineGroup 
Set mySG = Range("$A$1:$A$4").SparklineGroups.Add(Type:=xlSparkColumn, SourceData:= _ 
 "Sheet2!B1:E4") 
 
mySG.SeriesColor.Color = RGB(255, 0, 0)
```

## See also

[Excel Object Model Reference](./overview/Excel/object-model.md)

