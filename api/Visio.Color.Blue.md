---
title: Color.Blue property (Visio)
keywords: vis_sdr.chm12213145
f1_keywords:
- vis_sdr.chm12213145
ms.prod: visio
api_name:
- Visio.Color.Blue
ms.assetid: 7291912d-3521-5081-0e9d-4ce1bf1cccda
ms.date: 06/08/2017
localization_priority: Normal
---


# Color.Blue property (Visio)

Gets or sets the intensity of the blue component of a  **Color** object. Read/write.


## Syntax

_expression_.**Blue**

_expression_ A variable that represents a **[Color](Visio.Color.md)** object.


## Return value

Integer


## Remarks

The  **Blue** property can be a value from 0 to 255.

A color is represented by red, green, and blue components. It also has a flag that indicates how the color is to be used. These correspond to members of the Microsoft Windows  **PALETTEENTRY** data structure. For details, search for "PALETTEENTRY" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]