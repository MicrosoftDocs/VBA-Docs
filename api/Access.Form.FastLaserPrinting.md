---
title: Form.FastLaserPrinting property (Access)
keywords: vbaac10.chm13392
f1_keywords:
- vbaac10.chm13392
ms.prod: access
api_name:
- Access.Form.FastLaserPrinting
ms.assetid: a64775e5-174d-0349-d3f3-0009798d6462
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.FastLaserPrinting property (Access)

You can use the **FastLaserPrinting** property to specify whether lines and rectangles are replaced by text character lines, similar to the underscore ( _ ) and vertical bar ( | ) characters, when you print a form by using most laser printers. Replacing lines and rectangles with text character lines can make printing much faster. Read/write **Boolean**.


## Syntax

_expression_.**FastLaserPrinting**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **FastLaserPrinting** property affects any line or rectangle on a form or report, including controls that have these shapes (for example, a border around a text box).

This property has no effect on PostScript printers, dot-matrix printers, or earlier versions of laser printers that don't support text character lines.

When this property is set to **True** and the form or report being printed has overlapping rectangles or lines, the rectangles or lines on top don't erase the rectangles or lines that they overlap. If you require overlapping graphics on your report, set the **FastLaserPrinting** property to **False**.


## Example

The following example shows how to set the **FastLaserPrinting** property for the **Invoice** report while in report Design view.

```vb
DoCmd.OpenReport "Invoice", acDesign 
Reports!Invoice.FastLaserPrinting = True 
DoCmd.Close acReport, "Invoice", acSaveYes
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]