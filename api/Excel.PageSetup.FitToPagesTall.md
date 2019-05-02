---
title: PageSetup.FitToPagesTall property (Excel)
keywords: vbaxl10.chm473082
f1_keywords:
- vbaxl10.chm473082
ms.prod: excel
api_name:
- Excel.PageSetup.FitToPagesTall
ms.assetid: 1a0141cb-a665-caf5-6bd6-b037f65486dc
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.FitToPagesTall property (Excel)

Returns or sets the number of pages tall that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write **Variant**.


## Syntax

_expression_.**FitToPagesTall**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

If this property is **False**, Microsoft Excel scales the worksheet according to the **[FitToPagesWide](Excel.PageSetup.FitToPagesWide.md)** property.

If the **[Zoom](Excel.PageSetup.Zoom.md)** property is **True**, the **FitToPagesTall** property is ignored.


## Example

This example causes Microsoft Excel to print Sheet1 exactly one page tall and wide.

```vb
With Worksheets("Sheet1").PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
