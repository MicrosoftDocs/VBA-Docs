---
title: ApplicationSettings.TransitionsEnabled property (Visio)
keywords: vis_sdr.chm16262495
f1_keywords:
- vis_sdr.chm16262495
ms.prod: visio
api_name:
- Visio.ApplicationSettings.TransitionsEnabled
ms.assetid: af3b25b8-eee2-110f-9189-5133144d3a43
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.TransitionsEnabled property (Visio)

Determines whether Microsoft Visio uses an animated transition to show certain shape movements, such as re-layout of shapes. Read/write.


## Syntax

_expression_.**TransitionsEnabled**

_expression_ A variable that represents an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

 **Boolean**


## Remarks

The  **TransitionsEnabled** property setting also determines whether changes of view (for example, those initiated in the **Pan & Zoom** window) are animated.

The setting of the  **TransitionsEnabled** property corresponds to the setting of the **Enable transitions** check box under **Editing options** on the **Advanced** tab of the **Visio Options** dialog box. To open the **Visio Options** dialog box, click the **File** tab, click **Options**, and then click  **Advanced**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]