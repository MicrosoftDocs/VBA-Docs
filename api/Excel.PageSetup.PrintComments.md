---
title: PageSetup.PrintComments property (Excel)
keywords: vbaxl10.chm473104
f1_keywords:
- vbaxl10.chm473104
ms.prod: excel
api_name:
- Excel.PageSetup.PrintComments
ms.assetid: 1f479032-ca02-982f-5877-83c776ce2611
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintComments property (Excel)

Returns or sets the way comments are printed with the sheet. Read/write **[XlPrintLocation](Excel.XlPrintLocation.md)**.


## Syntax

_expression_.**PrintComments**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example causes comments to be printed as end notes when worksheet one is printed.

```vb
Worksheets(1).PageSetup.PrintComments = xlPrintSheetEnd
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]