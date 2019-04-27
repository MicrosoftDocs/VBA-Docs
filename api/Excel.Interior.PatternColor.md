---
title: Interior.PatternColor property (Excel)
keywords: vbaxl10.chm551077
f1_keywords:
- vbaxl10.chm551077
ms.prod: excel
api_name:
- Excel.Interior.PatternColor
ms.assetid: 44d3e506-56a4-e021-4b7c-452169a6dbf2
ms.date: 04/27/2019
localization_priority: Normal
---


# Interior.PatternColor property (Excel)

Returns or sets the color of the interior pattern as an **[RGB](../Language/Reference/User-Interface-Help/rgb-function.md)** value. Read/write **Variant**.


## Syntax

_expression_.**PatternColor**

_expression_ A variable that represents an **[Interior](excel.interior(object).md)** object.


## Example

This example sets the color of the interior pattern for rectangle one on Sheet1.

```vb
With Worksheets("Sheet1").Rectangles(1).Interior 
 .Pattern = xlGrid 
 .PatternColor = RGB(255,0,0) 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]