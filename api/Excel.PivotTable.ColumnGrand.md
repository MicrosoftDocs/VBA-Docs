---
title: PivotTable.ColumnGrand property (Excel)
keywords: vbaxl10.chm235075
f1_keywords:
- vbaxl10.chm235075
ms.prod: excel
api_name:
- Excel.PivotTable.ColumnGrand
ms.assetid: aa012e55-c944-22f1-13da-7ad76ae72c5b
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ColumnGrand property (Excel)

**True** if the PivotTable report shows grand totals for columns. Read/write **Boolean**.


## Syntax

_expression_.**ColumnGrand**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example sets the PivotTable report to show grand totals for columns.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.ColumnGrand = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]