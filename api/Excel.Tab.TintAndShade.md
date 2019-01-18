---
title: Tab.TintAndShade property (Excel)
keywords: vbaxl10.chm723076
f1_keywords:
- vbaxl10.chm723076
ms.prod: excel
api_name:
- Excel.Tab.TintAndShade
ms.assetid: be8ee335-fcf0-2091-89c1-973f54aba2c4
ms.date: 06/08/2017
localization_priority: Normal
---


# Tab.TintAndShade property (Excel)

Returns or sets a  **Single** that lightens or darkens a color.


## Syntax

_expression_. `TintAndShade`

_expression_ A variable that represents a [Tab](./Excel.Tab.md) object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1, is not recommended. Excel will correct the value internally to some value that falls with the range of valid values. This property works for both theme colors and nontheme colors.


## See also


[Tab Object](Excel.Tab.md)

