---
title: PageSetup.FitToPagesWide property (Excel)
keywords: vbaxl10.chm473083
f1_keywords:
- vbaxl10.chm473083
ms.prod: excel
api_name:
- Excel.PageSetup.FitToPagesWide
ms.assetid: 162bd2d2-35fa-8133-ab1c-27dcfc173317
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.FitToPagesWide property (Excel)

Returns or sets the number of pages wide that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write **Variant**.


## Syntax

_expression_.**FitToPagesWide**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

If this property is **False**, Microsoft Excel scales the worksheet according to the **[FitToPagesTall](Excel.PageSetup.FitToPagesTall.md)** property.

If the **[Zoom](Excel.PageSetup.Zoom.md)** property is **True**, the **FitToPagesWide** property is ignored.


## Example

This example causes Microsoft Excel to print Sheet1 exactly one page wide and tall.

```vb
With Worksheets("Sheet1").PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
