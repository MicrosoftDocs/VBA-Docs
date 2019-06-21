---
title: MSGWrap.PostTime property (Visio)
keywords: vis_sdr.chm16150770
f1_keywords:
- vis_sdr.chm16150770
ms.prod: visio
api_name:
- Visio.MSGWrap.PostTime
ms.assetid: e43c865b-eca8-22c7-de8e-1c6ec3f53348
ms.date: 06/08/2017
localization_priority: Normal
---


# MSGWrap.PostTime property (Visio)

Gets or sets the  **Time** member of the **MSG** structure being wrapped. Read/write.


## Syntax

_expression_.**PostTime**

_expression_ A variable that represents an **[MSGWrap](Visio.MSGWrap.md)** object.


## Return value

Long


## Remarks

The  **PostTime** property corresponds to the **Time** member of the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]