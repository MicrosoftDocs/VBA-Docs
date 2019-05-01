---
title: PageSetup.FooterMargin property (Excel)
keywords: vbaxl10.chm473084
f1_keywords:
- vbaxl10.chm473084
ms.prod: excel
api_name:
- Excel.PageSetup.FooterMargin
ms.assetid: b6ec4b9c-c828-e6fe-2a65-ccddd1b05c30
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.FooterMargin property (Excel)

Returns or sets the distance from the bottom of the page to the footer, in points. Read/write  **Double**.


## Syntax

_expression_.**FooterMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example sets the footer margin of Sheet1 to 0.5 inch.


```vb
Worksheets("Sheet1").PageSetup.FooterMargin = _ 
 Application.InchesToPoints(0.5)
```


## See also


[PageSetup Object](Excel.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]