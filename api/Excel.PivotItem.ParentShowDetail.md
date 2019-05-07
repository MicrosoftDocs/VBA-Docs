---
title: PivotItem.ParentShowDetail property (Excel)
keywords: vbaxl10.chm246080
f1_keywords:
- vbaxl10.chm246080
ms.prod: excel
api_name:
- Excel.PivotItem.ParentShowDetail
ms.assetid: 7700aa5c-e90a-864f-b907-a84656ecdaaa
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItem.ParentShowDetail property (Excel)

**True** if the specified item is showing because one of its parents is showing detail. **False** if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. Read-only **Boolean**.


## Syntax

_expression_.**ParentShowDetail**

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Remarks

This property isn't available for OLAP data sources.


## Example

This example displays a message if the item that contains the active cell is visible because its parent item is showing detail.

```vb
Worksheets("Sheet1").Activate 
Set pvtItem = ActiveCell.PivotItem 
If pvtItem.ParentShowDetail = True Then 
 MsgBox "Parent item is showing detail" 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]