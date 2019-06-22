---
title: MSGWrap.Message property (Visio)
keywords: vis_sdr.chm16150730
f1_keywords:
- vis_sdr.chm16150730
ms.prod: visio
api_name:
- Visio.MSGWrap.Message
ms.assetid: ae780612-a017-93b8-1c39-abe8097dfbf2
ms.date: 06/08/2017
localization_priority: Normal
---


# MSGWrap.message property (Visio)

Gets or sets the **Message** member of the **MSG** structure being wrapped. Read/write.


## Syntax

_expression_. `Message`

_expression_ A variable that represents an **[MSGWrap](Visio.MSGWrap.md)** object.


## Return value

Long


## Remarks

The **Message** property corresponds to the **Message** member of the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

The **OnKeystrokeMessageForAddon** event fires for messages in the following range:



|WM_KEYDOWN|0x0100|
|WM_KEYUP |0x0101 |
|WM_CHAR |0x0102 |
|WM_DEADCHAR |0x0103 |
|WM_SYSKEYDOWN |0x0104 |
|WM_SYSKEYUP |0x0105 |
|WM_SYSCHAR |0x0106 |
|WM_SYSDEADCHAR |0x0107 |



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]