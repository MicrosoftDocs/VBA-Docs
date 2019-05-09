---
title: PivotTable.RefreshName property (Excel)
keywords: vbaxl10.chm235091
f1_keywords:
- vbaxl10.chm235091
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshName
ms.assetid: 488d5e0c-61f9-0c85-ac1b-16dc98360bb4
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RefreshName property (Excel)

Returns the name of the person who last refreshed the PivotTable report data. Read-only **String**.


## Syntax

_expression_.**RefreshName**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For OLAP data sources, this property is updated after each query.


## Example

This example displays the name of the person who last refreshed the PivotTable report.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox "The data was last refreshed by " & pvtTable.RefreshName
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]