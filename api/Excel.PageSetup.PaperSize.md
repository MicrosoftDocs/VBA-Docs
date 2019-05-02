---
title: PageSetup.PaperSize property (Excel)
keywords: vbaxl10.chm473091
f1_keywords:
- vbaxl10.chm473091
ms.prod: excel
api_name:
- Excel.PageSetup.PaperSize
ms.assetid: 7c26e996-8399-31b4-8e53-772de8bf8eb2
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PaperSize property (Excel)

Returns or sets the size of the paper. Read/write **[XlPaperSize](Excel.XlPaperSize.md)**.


## Syntax

_expression_.**PaperSize**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Some printers may not support all the paper sizes listed on the **XlPaperSize** enumeration.


## Example

This example sets the paper size to legal for Sheet1.

```vb
Worksheets("Sheet1").PageSetup.PaperSize = xlPaperLegal
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
