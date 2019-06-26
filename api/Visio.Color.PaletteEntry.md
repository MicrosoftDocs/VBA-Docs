---
title: Color.PaletteEntry property (Visio)
keywords: vis_sdr.chm12214005
f1_keywords:
- vis_sdr.chm12214005
ms.prod: visio
api_name:
- Visio.Color.PaletteEntry
ms.assetid: 4a761fc2-6696-dc44-6d23-ff630a76bdd4
ms.date: 06/08/2017
localization_priority: Normal
---


# Color.PaletteEntry property (Visio)

Gets or sets the red, green, blue, and flags components of a color. Read/write.


## Syntax

_expression_.**PaletteEntry**

_expression_ A variable that represents a **[Color](Visio.Color.md)** object.


## Return value

Long


## Remarks

A color is represented by 1-byte red, green, and blue components. It also has a 1-byte flags field that indicates how you use the color. These correspond to members of the Windows  **PALETTEENTRY** data structure. For details, search for "PALETTEENTRY" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

The value passed is four tightly packed BYTE fields. The correspondence between the  **PaletteEntry** property and red, green, blue, and flags values is:




```vb
    paletteentry == r+256(b+256(g+256f))
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]