---
title: FormatConditions object (Excel)
keywords: vbaxl10.chm509072
f1_keywords:
- vbaxl10.chm509072
ms.prod: excel
api_name:
- Excel.FormatConditions
ms.assetid: 2486d4b4-605c-76d8-132a-694c0c600a81
ms.date: 03/30/2019
localization_priority: Normal
---


# FormatConditions object (Excel)

Represents the collection of conditional formats for a single range.


## Remarks

The **FormatConditions** collection can contain multiple conditional formats. Each format is represented by a **[FormatCondition](Excel.FormatCondition.md)** object.

Use the **[FormatConditions](Excel.Range.FormatConditions.md)** property to return a **FormatConditions** object. Use the **Add** method to create a new conditional format, and use the **[Modify](Excel.FormatCondition.Modify.md)** method of the **FormatCondition** object to change an existing conditional format.


## Example

The following example adds a conditional format to cells E1:E10.

```vb
With Worksheets(1).Range("e1:e10").FormatConditions _ 
 .Add(xlCellValue, xlGreater, "=$a$1") 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```


## Methods

- [Add](Excel.FormatConditions.Add.md)
- [AddAboveAverage](Excel.FormatConditions.AddAboveAverage.md)
- [AddColorScale](Excel.FormatConditions.AddColorScale.md)
- [AddDataBar](Excel.FormatConditions.AddDatabar.md)
- [AddIconSetCondition](Excel.FormatConditions.AddIconSetCondition.md)
- [AddTop10](Excel.FormatConditions.AddTop10.md)
- [AddUniqueValues](Excel.FormatConditions.AddUniqueValues.md)
- [Delete](Excel.FormatConditions.Delete.md)
- [Item](Excel.FormatConditions.Item.md)

## Properties

- [Application](Excel.FormatConditions.Application.md)
- [Count](Excel.FormatConditions.Count.md)
- [Creator](Excel.FormatConditions.Creator.md)
- [Parent](Excel.FormatConditions.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
