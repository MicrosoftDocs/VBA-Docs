---
title: PivotTable.RefreshDate property (Excel)
keywords: vbaxl10.chm235090
f1_keywords:
- vbaxl10.chm235090
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshDate
ms.assetid: 7c1a29c2-749e-98f8-ae14-eb2fa3ab2bb1
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RefreshDate property (Excel)

Returns the date on which the PivotTable report was last refreshed. Read-only **Date**.


## Syntax

_expression_.**RefreshDate**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For OLAP data sources, this property is updated after each query.


## Example

This example displays the date on which the PivotTable report was last refreshed.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
dateString = Format(pvtTable.RefreshDate, "Long Date") 
MsgBox "The data was last refreshed on " & dateString
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]