---
title: Page.SetThemeVariant method (Visio)
ms.prod: visio
ms.assetid: 8393a95f-83ca-0efa-d987-ae498bfe5e9d
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.SetThemeVariant method (Visio)

Sets the color, style, and optionally the embellishment of the variant of the theme applied to the specified page.


## Syntax

_expression_.**SetThemeVariant** (_variantColor_, _variantStyle_, _embellishment_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _variantColor_|Required|INT16|The index of the color variant to apply. Possible values are from 0 to 3.|
| _variantStyle_|Required|INT16|The index of the style variant to apply. Possible values are from 0 to 3.|
| _embellishment_|Optional|INT16|The index of the embellishment to apply. Possible values are from 1, for low, to 3, for high.|

## Return value

**VOID**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]