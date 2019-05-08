---
title: PivotTable.PrintTitles property (Excel)
keywords: vbaxl10.chm235131
f1_keywords:
- vbaxl10.chm235131
ms.prod: excel
api_name:
- Excel.PivotTable.PrintTitles
ms.assetid: a8138146-bfe9-1af9-c101-0c095c4a91a5
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PrintTitles property (Excel)

**True** if the print titles for the worksheet are set based on the PivotTable report. **False** if the print titles for the worksheet are used. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**PrintTitles**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

The row print titles are set to the rows that contain the PivotTable report's column field items. The column print titles are set to the columns that contain the row items.


## Example

This example specifies that the print title set for the worksheet is printed when the fourth PivotTable report on the active worksheet is printed.

```vb
ActiveSheet.PivotTables("PivotTable4").PrintTitles = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]