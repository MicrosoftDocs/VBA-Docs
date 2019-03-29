---
title: PivotFields object (Excel)
keywords: vbaxl10.chm241072
f1_keywords:
- vbaxl10.chm241072
ms.prod: excel
api_name:
- Excel.PivotFields
ms.assetid: 018d4cea-09ea-d4be-baef-5fd55062935b
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotFields object (Excel)

A collection of all the **[PivotField](Excel.PivotField.md)** objects in a PivotTable report.


## Remarks

In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following properties are available:

- **[ColumnFields](Excel.PivotTable.ColumnFields.md)** property    
- **[DataFields](Excel.PivotTable.DataFields.md)** property    
- **[HiddenFields](Excel.PivotTable.HiddenFields.md)** property   
- **[PageFields](Excel.PivotTable.PageFields.md)** property   
- **[RowFields](Excel.PivotTable.RowFields.md)** property   
- **[VisibleFields](Excel.PivotTable.VisibleFields.md)** property
    

## Example

Use the **[PivotFields](Excel.PivotTable.PivotFields.md)** method of the **PivotTable** object to return the **PivotFields** collection. 

The following example enumerates the field names in the first PivotTable report on Sheet3.

```vb
With Worksheets("sheet3").PivotTables(1) 
 For i = 1 To .PivotFields.Count 
 MsgBox .PivotFields(i).Name 
 Next 
End With
```

<br/>

Use **PivotFields** (_index_), where _index_ is the field name or index number, to return a single **PivotField** object. The following example makes the Year field a row field in the first PivotTable report on Sheet3.

```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods

- [Item](Excel.PivotFields.Item.md)

## Properties

- [Application](Excel.PivotFields.Application.md)
- [Count](Excel.PivotFields.Count.md)
- [Creator](Excel.PivotFields.Creator.md)
- [Parent](Excel.PivotFields.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]