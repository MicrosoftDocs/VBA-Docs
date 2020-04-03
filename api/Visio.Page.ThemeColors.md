---
title: Page.ThemeColors property (Visio)
keywords: vis_sdr.chm10960180
f1_keywords:
- vis_sdr.chm10960180
ms.prod: visio
api_name:
- Visio.Page.ThemeColors
ms.assetid: a3f4bc4e-3dbb-9d50-9d71-f77b39ec0ac3
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ThemeColors property (Visio)

Gets or sets the current theme colors for the page. Read/write.


## Syntax

_expression_. `ThemeColors`

 _expression_ An expression that returns a **[Page](Visio.Page.md)** object.


## Return value

Variant


## Remarks

You can set the  **ThemeColors** property value to any one of the following:




- The name or universal name of the theme color (strings)
    
- An enumerated value from the  **[VisThemeColors](Visio.visthemecolors.md)** enumeration
    
- A  **Master** object of type **visTypeThemeColors**
    


The  **ThemeColors** property always returns the universal name of the current theme colors.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]