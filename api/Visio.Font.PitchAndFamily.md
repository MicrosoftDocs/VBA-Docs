---
title: Font.PitchAndFamily property (Visio)
keywords: vis_sdr.chm12014085
f1_keywords:
- vis_sdr.chm12014085
ms.prod: visio
api_name:
- Visio.Font.PitchAndFamily
ms.assetid: 1902eb17-9be5-7337-bfdc-7804c66555ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.PitchAndFamily property (Visio)

Returns the pitch and family code for a  **Font** object. Read-only.


## Syntax

_expression_. `PitchAndFamily`

_expression_ A variable that represents a **[Font](Visio.Font.md)** object.


## Return value

Integer


## Remarks

The possible values of the  **PitchAndFamily** property correspond to those of the **lfPitchAndFamily** member of the Windows **LOGFONT** data structure. For details, search for "LOGFONT" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]