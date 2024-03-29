---
title: Font.TintAndShade property (Excel)
keywords: vbaxl10.chm559088
f1_keywords:
- vbaxl10.chm559088
api_name:
- Excel.Font.TintAndShade
ms.assetid: 0b890357-fb55-ac43-ecf0-f7d84ce0f248
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Font.TintAndShade property (Excel)

Returns or sets a **Single** that lightens or darkens a color.


## Syntax

_expression_.**TintAndShade**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and non-theme colors.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]