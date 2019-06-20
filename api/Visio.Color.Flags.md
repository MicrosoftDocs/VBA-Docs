---
title: Color.Flags property (Visio)
keywords: vis_sdr.chm12213540
f1_keywords:
- vis_sdr.chm12213540
ms.prod: visio
api_name:
- Visio.Color.Flags
ms.assetid: 61289973-af74-4eca-f4ac-becb9ca35ed4
ms.date: 06/08/2017
localization_priority: Normal
---


# Color.Flags property (Visio)

Gets or sets the flags that specify how you use a  **Color** object. Read/write.


## Syntax

_expression_.**Flags**

_expression_ A variable that represents a **[Color](Visio.Color.md)** object.


## Return value

Integer


## Remarks

The  **Flags** property of a **Color** object corresponds to the **peFlags** member of a Microsoft Windows **PALETTEENTRY** data structure. For details, search for "PALETTEENTRY" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]