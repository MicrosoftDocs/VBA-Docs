---
title: PivotTable.RepeatItemsOnEachPrintedPage property (Excel)
keywords: vbaxl10.chm235135
f1_keywords:
- vbaxl10.chm235135
ms.prod: excel
api_name:
- Excel.PivotTable.RepeatItemsOnEachPrintedPage
ms.assetid: 96e5e2d8-44ff-8d6f-6bba-f009dbc769a7
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RepeatItemsOnEachPrintedPage property (Excel)

**True** if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. **False** if labels are printed only on the first page. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**RepeatItemsOnEachPrintedPage**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Microsoft Excel prints row and column labels in place of any print titles set for the worksheet. Use the **[PrintTitles](Excel.PivotTable.PrintTitles.md)** property to determine whether print titles are set for the PivotTable report.


## Example

This example sets Microsoft Excel to repeat the labels on each page when the fourth PivotTable report on the active worksheet is printed.

```vb
ActiveSheet.PivotTables("PivotTable4") _ 
 .RepeatItemsOnEachPrintedPage = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]