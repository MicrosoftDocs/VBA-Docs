---
title: PivotItems object (Excel)
keywords: vbaxl10.chm247072
f1_keywords:
- vbaxl10.chm247072
ms.prod: excel
api_name:
- Excel.PivotItems
ms.assetid: df47021a-2b06-fa10-5712-58956c7ffe07
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotItems object (Excel)

A collection of all the **[PivotItem](Excel.PivotItem.md)** objects in a PivotTable field.


## Remarks

The items are the individual data entries in a field category.


## Example

Use the **[PivotItems](Excel.PivotField.PivotItems.md)** method of the **PivotField** object to return the **PivotItems** collection. 

The following example creates an enumerated list of field names and the items contained in those fields for the first PivotTable report on Sheet4.

```vb
Worksheets("sheet4").Activate 
With Worksheets("sheet3").PivotTables(1) 
 c = 1 
 For i = 1 To .PivotFields.Count 
 r = 1 
 Cells(r, c) = .PivotFields(i).Name 
 r = r + 1 
 For x = 1 To .PivotFields(i).PivotItems.Count 
 Cells(r, c) = .PivotFields(i).PivotItems(x).Name 
 r = r + 1 
 Next 
 c = c + 1 
 Next 
End With
```

<br/>

Use **PivotItems** (_index_), where _index_ is the item index number or name, to return a single **PivotItem** object. The following example hides all entries in the first PivotTable report on Sheet3 that contain "1998" in the Year field.

```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").PivotItems("1998").Visible = False
```


## Methods

- [Add](Excel.PivotItems.Add.md)
- [Item](Excel.PivotItems.Item.md)

## Properties

- [Application](Excel.PivotItems.Application.md)
- [Count](Excel.PivotItems.Count.md)
- [Creator](Excel.PivotItems.Creator.md)
- [Parent](Excel.PivotItems.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]