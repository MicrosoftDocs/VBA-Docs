---
title: Border.TintAndShade property (Excel)
keywords: vbaxl10.chm547078
f1_keywords:
- vbaxl10.chm547078
ms.prod: excel
api_name:
- Excel.Border.TintAndShade
ms.assetid: 3ec15506-3ba6-a173-a11b-d17448fcdb1b
ms.date: 03/07/2019
localization_priority: Normal
---


# Border.TintAndShade property (Excel)

Returns or sets a **Single** that lightens or darkens a color.


## Syntax

_expression_.**TintAndShade**

_expression_ A variable that represents a **[Border](Excel.Border(object).md)** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in this run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.

> [!IMPORTANT] 
> Note that the visual properties of a **Border** object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. For an example, see the **[Border](excel.border(object).md)** object.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]