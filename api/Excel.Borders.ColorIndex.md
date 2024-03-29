---
title: Borders.ColorIndex property (Excel)
keywords: vbaxl10.chm181074
f1_keywords:
- vbaxl10.chm181074
api_name:
- Excel.Borders.ColorIndex
ms.assetid: fe0a7b5e-254d-c773-88cc-70728db44840
ms.date: 04/13/2019
ms.localizationpriority: medium
---


# Borders.ColorIndex property (Excel)

Returns or sets a **Variant** value that represents the color of all four borders.


## Syntax

_expression_.**ColorIndex**

_expression_ A variable that represents a **[Borders](Excel.Borders.md)** object.


## Remarks

This property returns **Null** if all four borders aren't the same color.

The color is specified as an index value into the current color palette, or as one of the following **[XlColorIndex](Excel.XlColorIndex.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]