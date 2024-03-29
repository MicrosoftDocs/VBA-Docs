---
title: Interior.TintAndShade property (Excel)
keywords: vbaxl10.chm551080
f1_keywords:
- vbaxl10.chm551080
api_name:
- Excel.Interior.TintAndShade
ms.assetid: 45b12e93-1a6d-b5a3-b31d-4b41d87f3f73
ms.date: 04/27/2019
ms.localizationpriority: medium
---


# Interior.TintAndShade property (Excel)

Returns or sets a **Single** that lightens or darkens a color.


## Syntax

_expression_.**TintAndShade**

_expression_ A variable that represents an **[Interior](excel.interior(object).md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
