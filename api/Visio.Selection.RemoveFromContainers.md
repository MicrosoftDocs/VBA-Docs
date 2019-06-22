---
title: Selection.RemoveFromContainers method (Visio)
keywords: vis_sdr.chm11162220
f1_keywords:
- vis_sdr.chm11162220
ms.prod: visio
api_name:
- Visio.Selection.RemoveFromContainers
ms.assetid: d1ed1360-3caa-3e03-98ef-84f4bd52a035
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.RemoveFromContainers method (Visio)

Removes the selection of shapes from all lists and containers of which the selection is a member.


## Syntax

_expression_. `RemoveFromContainers`

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Return value

 **Nothing**


## Remarks

When you call the  **RemoveFromContainers** method, Microsoft Visio uses the setting of the **[ContainerProperties.ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property for each container to determine how the container resizes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]