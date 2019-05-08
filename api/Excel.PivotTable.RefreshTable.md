---
title: PivotTable.RefreshTable method (Excel)
keywords: vbaxl10.chm235092
f1_keywords:
- vbaxl10.chm235092
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshTable
ms.assetid: 778743e3-c53a-23e3-73c6-c18339cd1ac2
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RefreshTable method (Excel)

Refreshes the PivotTable report from the source data. Returns **True** if it's successful.


## Syntax

_expression_.**RefreshTable**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

Boolean


## Example

This example refreshes the PivotTable report.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.RefreshTable
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
