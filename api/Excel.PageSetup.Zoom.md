---
title: PageSetup.Zoom property (Excel)
keywords: vbaxl10.chm473103
f1_keywords:
- vbaxl10.chm473103
ms.prod: excel
api_name:
- Excel.PageSetup.Zoom
ms.assetid: 3deebce5-8605-c549-371c-033848073ffe
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.Zoom property (Excel)

Returns or sets a **Variant** value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.


## Syntax

_expression_.**Zoom**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

This property applies only to worksheets.

If this property is **False**, the **[FitToPagesWide](Excel.PageSetup.FitToPagesWide.md)** and **[FitToPagesTall](Excel.PageSetup.FitToPagesTall.md)** properties control how the worksheet is scaled.

All scaling retains the aspect ratio of the original document.


## Example

This example scales Sheet1 by 150 percent when the worksheet is printed.

```vb
Worksheets("Sheet1").PageSetup.Zoom = 150
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
