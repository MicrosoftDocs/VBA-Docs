---
title: ColorFormat.TintAndShade property (Excel)
keywords: vbaxl10.chm105006
f1_keywords:
- vbaxl10.chm105006
ms.prod: excel
api_name:
- Excel.ColorFormat.TintAndShade
ms.assetid: b548b2ad-da3d-0d02-249e-2ab37271a5c6
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorFormat.TintAndShade property (Excel)

Returns or sets a **Single** that lightens or darkens a color.


## Syntax

_expression_.**TintAndShade**

_expression_ A variable that represents a **[ColorFormat](Excel.ColorFormat.md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]