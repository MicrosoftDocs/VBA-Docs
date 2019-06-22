---
title: Window.MergeCaption property (Visio)
keywords: vis_sdr.chm11650940
f1_keywords:
- vis_sdr.chm11650940
ms.prod: visio
api_name:
- Visio.Window.MergeCaption
ms.assetid: 19461100-0242-28b1-60bc-9b7f2da3af02
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.MergeCaption property (Visio)

Gets or sets the abbreviated caption that appears on the page tab when the window is merged with other windows. Read/write.


## Syntax

_expression_. `MergeCaption`

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Return value

String


## Remarks

The  **MergeCaption** property applies only to anchored windows. If the **Window** object is an MDI frame window, Microsoft Visio raises an exception.

Use the  **Type** property to determine window type.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]