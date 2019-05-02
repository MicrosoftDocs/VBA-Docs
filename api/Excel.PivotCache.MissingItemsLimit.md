---
title: PivotCache.MissingItemsLimit property (Excel)
keywords: vbaxl10.chm227102
f1_keywords:
- vbaxl10.chm227102
ms.prod: excel
api_name:
- Excel.PivotCache.MissingItemsLimit
ms.assetid: ff15a86c-b57f-ed55-bbfa-74e1c5ce753c
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.MissingItemsLimit property (Excel)

Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write **[XlPivotTableMissingItems](Excel.XlPivotTableMissingItems.md)**.


## Syntax

_expression_.**MissingItemsLimit**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

This property can be set to a value between 0 and 32,500. If an integer less than zero is specified, this is equivalent to specifying **xlMissingItemsDefault**. Integers greater than 32,500 can be specified but will have the same effect as specifying **xlMissingItemsMax**.

The **MissingItemsLimit** property only works for non-OLAP PivotTables; otherwise, a run-time error can occur.


## Example

This example determines the maximum quantity of unique items per field and notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckMissingItemsList() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the maximum number of unique items allowed per PivotField and notify the user. 
 Select Case pvtCache.MissingItemsLimit 
 Case xlMissingItemsDefault 
 MsgBox "The default value of unique items per PivotField is allowed." 
 Case xlMissingItemsMax 
 MsgBox "The maximum value of unique items per PivotField is allowed." 
 Case xlMissingItemsNone 
 MsgBox "No unique items per PivotField are allowed." 
 End Select 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]