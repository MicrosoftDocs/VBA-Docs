---
title: Page.ThemeEffects property (Visio)
keywords: vis_sdr.chm10960185
f1_keywords:
- vis_sdr.chm10960185
ms.prod: visio
api_name:
- Visio.Page.ThemeEffects
ms.assetid: 566ee9aa-9c45-e53b-2634-c666565e6fbb
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ThemeEffects property (Visio)

Gets or sets the current theme effect for the page. Read/write.


## Syntax

_expression_.**ThemeEffects**

_expression_ An expression that returns a **[Page](Visio.Page.md)** object.


## Return value

Variant


## Remarks

You can set the **ThemeEffects** property value to any one of the following:

- The name or universal name of the theme effect (strings)
    
- An enumerated value from the **[VisThemeEffects](Visio.visthemeeffects.md)** enumeration
    
- A **Master** object of type **visTypeThemeEffects**
    

The **ThemeEffects** property always returns the universal name of the current theme effect.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]