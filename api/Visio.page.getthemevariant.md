---
title: Page.GetThemeVariant method (Visio)
ms.prod: visio
ms.assetid: 40c2be31-fdb0-68ee-a129-2788b1b17c82
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.GetThemeVariant method (Visio)

Returns the color, style, and embellishment, if any, of the variant of the theme applied to the specified page.


## Syntax

_expression_.**GetThemeVariant** (_pVariantColor_, _pVariantStyle_, _pEmbellishment_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pVariantColor_|Required|INT16|The index of the color variant applied. Possible values are from 0 to 3. Out parameter.|
| _pVariantStyle_|Required|INT16|The index of the style variant applied. Possible values are from 0 to 3. Out parameter.|
| _pEmbellishment_|Optional|INT16|The index of the embellishment applied, if any. Possible values are from 1, for low, to 3, for high. Out parameter.|

## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]