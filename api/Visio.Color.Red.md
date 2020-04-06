---
title: Color.Red property (Visio)
keywords: vis_sdr.chm12214190
f1_keywords:
- vis_sdr.chm12214190
ms.prod: visio
api_name:
- Visio.Color.Red
ms.assetid: aeb7a499-710d-de11-37c6-673aac54f27d
ms.date: 06/08/2017
localization_priority: Normal
---


# Color.Red property (Visio)

Gets or sets the intensity of the red component of a  **Color** object. Read/write.


## Syntax

_expression_.**Red**

_expression_ A variable that represents a **[Color](Visio.Color.md)** object.


## Return value

Integer


## Remarks

The  **Red** property can be a value from zero (0) to 255.

A color is represented by red, green, and blue components. It also has flags that indicate how the color is to be used. These correspond to members of the Microsoft Windows  **PALETTEENTRY** data structure. For details, search for "PALETTEENTRY" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]