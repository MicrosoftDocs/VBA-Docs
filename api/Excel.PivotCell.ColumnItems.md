---
title: PivotCell.ColumnItems property (Excel)
keywords: vbaxl10.chm692079
f1_keywords:
- vbaxl10.chm692079
ms.prod: excel
api_name:
- Excel.PivotCell.ColumnItems
ms.assetid: 66936e2f-740e-e3de-5d20-47885bee9691
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.ColumnItems property (Excel)

Returns a **[PivotItemList](Excel.PivotItemList.md)** collection that corresponds to the items on the column axis that represent the selected range.


## Syntax

_expression_.**ColumnItems**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Example

This example determines if the data item in cell B5 is under the Inventory item in the first column field, and notifies the user. The example assumes that a PivotTable exists on the active worksheet and that column B contains a column field of the PivotTable.

```vb
Sub CheckColumnItems() 
 
 ' Determine if there is a match between the item and column field. 
 If Application.Range("B5").PivotCell.ColumnItems.Item(1) = "Inventory" Then 
 MsgBox "Item in B5 is a member of the 'Inventory' column field." 
 Else 
 MsgBox "Item in B5 is not a member of the 'Inventory' column field." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]