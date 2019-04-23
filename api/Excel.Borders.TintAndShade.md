---
title: Borders.TintAndShade property (Excel)
keywords: vbaxl10.chm181083
f1_keywords:
- vbaxl10.chm181083
ms.prod: excel
api_name:
- Excel.Borders.TintAndShade
ms.assetid: 29c591bf-311e-5706-0222-1db144a92b77
ms.date: 04/13/2019
localization_priority: Normal
---


# Borders.TintAndShade property (Excel)

Returns or sets a **Single** that lightens or darkens a color.


## Syntax

_expression_.**TintAndShade**

_expression_ A variable that represents a **[Borders](Excel.Borders.md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]