---
title: Workbook.PivotCaches method (Excel)
keywords: vbaxl10.chm199124
f1_keywords:
- vbaxl10.chm199124
ms.prod: excel
api_name:
- Excel.Workbook.PivotCaches
ms.assetid: 0a2e7f10-c123-5c98-fb71-56868b9f8bde
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PivotCaches method (Excel)

Returns a **[PivotCaches](Excel.PivotCaches.md)** collection that represents all the PivotTable caches in the specified workbook. Read-only.


## Syntax

_expression_.**PivotCaches**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**PivotCaches**


## Example

This example causes the PivotTable cache to update automatically each time the workbook is opened.

```vb
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]