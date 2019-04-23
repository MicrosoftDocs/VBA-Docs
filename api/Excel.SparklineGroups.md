---
title: SparklineGroups object (Excel)
keywords: vbaxl10.chm868072
f1_keywords:
- vbaxl10.chm868072
ms.prod: excel
api_name:
- Excel.SparklineGroups
ms.assetid: 9bc6be34-fa2e-8652-ca92-fa9630b4d7a6
ms.date: 04/02/2019
localization_priority: Normal
---


# SparklineGroups object (Excel)

Represents a collection of sparkline groups.


## Remarks

The **SparklineGroups** object can contain multiple **[SparklineGroup](Excel.SparklineGroup.md)** objects.

Use the **[SparklineGroups](Excel.Range.SparklineGroups.md)** property of the **Range** object to return an existing **SparklineGroups** collection from its parent range.

Use the **Add** method to create a group of new sparklines.

Use the **Group** method to create a group of existing sparklines.


## Example

This example selects the range A1:A4 and groups the sparklines in that range. If the sparklines in the sparkline group are line sparklines, the markers are displayed in red.


```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Group Location := Range("A1") 
Selection.SparklineGroups.Item(1).Points.Markers.Visible = True 
Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 255
```

## Methods

- [Add](Excel.SparklineGroups.Add.md)
- [Clear](Excel.SparklineGroups.Clear.md)
- [ClearGroups](Excel.SparklineGroups.ClearGroups.md)
- [Group](Excel.SparklineGroups.Group.md)
- [Ungroup](Excel.SparklineGroups.Ungroup.md)

## Properties

- [Application](Excel.SparklineGroups.Application.md)
- [Count](Excel.SparklineGroups.Count.md)
- [Creator](Excel.SparklineGroups.Creator.md)
- [Item](Excel.SparklineGroups.Item.md)
- [Parent](Excel.SparklineGroups.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]