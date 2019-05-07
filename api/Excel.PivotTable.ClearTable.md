---
title: PivotTable.ClearTable method (Excel)
keywords: vbaxl10.chm235162
f1_keywords:
- vbaxl10.chm235162
ms.prod: excel
api_name:
- Excel.PivotTable.ClearTable
ms.assetid: 1279b0b8-3785-00b1-b91f-20e406ea1f2e
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ClearTable method (Excel)

The **ClearTable** method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.


## Syntax

_expression_.**ClearTable**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

The **ClearTable** function takes no arguments and is available for both relational and OLAP PivotTables.


## Example

The following example clears a PivotTable on the active worksheet.

```vb
ActiveSheet.PivotTables(1).ClearTable()
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
