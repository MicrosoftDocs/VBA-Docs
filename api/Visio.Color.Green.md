---
title: Color.Green property (Visio)
keywords: vis_sdr.chm12213610
f1_keywords:
- vis_sdr.chm12213610
ms.prod: visio
api_name:
- Visio.Color.Green
ms.assetid: 19d792e0-1fc7-e302-eb7d-8a80ad287a52
ms.date: 06/08/2017
localization_priority: Normal
---


# Color.Green property (Visio)

Gets or sets the intensity of the green component of a  **Color** object. Read/write.


## Syntax

_expression_.**Green**

_expression_ A variable that represents a **[Color](Visio.Color.md)** object.


## Return value

Integer


## Remarks

The  **Green** property can be a value from 0 to 255.

A color is represented by red, green, and blue components. It also has flags that indicate how the color is to be used. These correspond to members of the Microsoft Windows  **PALETTEENTRY** data structure. For details, search for "PALETTEENTRY" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]