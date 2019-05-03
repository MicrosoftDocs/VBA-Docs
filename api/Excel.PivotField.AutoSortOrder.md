---
title: PivotField.AutoSortOrder property (Excel)
keywords: vbaxl10.chm240113
f1_keywords:
- vbaxl10.chm240113
ms.prod: excel
api_name:
- Excel.PivotField.AutoSortOrder
ms.assetid: b2be072b-305a-5cdb-0602-368a67bed56f
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.AutoSortOrder property (Excel)

Returns the order used to sort the specified PivotTable field automatically. Can be one of the **[XlSortOrder](Excel.XlSortOrder.md)** constants. Read-only **Long**.


## Syntax

_expression_.**AutoSortOrder**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example displays a message box showing the **AutoSort** parameters for the Product field.

```vb
With Worksheets(1).PivotTables(1).PivotFields("product") 
 Select Case .AutoSortOrder 
 Case xlManual 
 aso = "manual" 
 Case xlAscending 
 aso = "ascending" 
 Case xlDescending 
 aso = "descending" 
 End Select 
 MsgBox " sorted in " & aso & _ 
 " order by " & .AutoSortField 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]