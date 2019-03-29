---
title: PivotItemList object (Excel)
keywords: vbaxl10.chm720072
f1_keywords:
- vbaxl10.chm720072
ms.prod: excel
api_name:
- Excel.PivotItemList
ms.assetid: 2b0fc8e5-6073-9cb1-2217-1e8715cddb1e
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotItemList object (Excel)

A collection of all the **[PivotItem](Excel.PivotItem.md)** objects in the specified PivotTable.


## Remarks

Each **PivotItem** represents an item in a PivotTable field.

Use the **[RowItems](Excel.PivotCell.RowItems.md)** or **[ColumnItems](Excel.PivotCell.ColumnItems.md)** property of the **PivotCell** object to return a **PivotItemList** collection.


## Example

After a **PivotItemList** collection is returned, you can use the **[Item](Excel.PivotItems.Item.md)** method of the **PivotItems** object to identify a particular **PivotItem** list. 

The following example displays the **PivotItem** list associated with cell B5 to the user. This example assumes that a PivotTable exists on the active worksheet.


```vb
Sub CheckPivotItemList() 
 
 ' Identify contents associated with PivotItemList. 
 MsgBox "Contents associated with cell B5: " & _ 
 Application.Range("B5").PivotCell.RowItems.Item(1) 
 
End Sub
```

## Methods

- [Item](Excel.PivotItemList.Item.md)

## Properties

- [Application](Excel.PivotItemList.Application.md)
- [Count](Excel.PivotItemList.Count.md)
- [Creator](Excel.PivotItemList.Creator.md)
- [Parent](Excel.PivotItemList.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]