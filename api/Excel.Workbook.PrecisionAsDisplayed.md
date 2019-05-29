---
title: Workbook.PrecisionAsDisplayed property (Excel)
keywords: vbaxl10.chm199126
f1_keywords:
- vbaxl10.chm199126
ms.prod: excel
api_name:
- Excel.Workbook.PrecisionAsDisplayed
ms.assetid: 4f0c8201-5b8d-5cb5-337c-944d2c7dd8d1
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PrecisionAsDisplayed property (Excel)

**True** if calculations in this workbook are done by using only the precision of the numbers as they're displayed. Read/write **Boolean**.


## Syntax

_expression_.**PrecisionAsDisplayed**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example causes calculations in the active workbook to use only the precision of the numbers as they're displayed.

```vb
ActiveWorkbook.PrecisionAsDisplayed = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]