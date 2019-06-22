---
title: Shape.RemoveFromContainers method (Visio)
keywords: vis_sdr.chm11262220
f1_keywords:
- vis_sdr.chm11262220
ms.prod: visio
api_name:
- Visio.Shape.RemoveFromContainers
ms.assetid: b9dbf604-01f0-675a-a0e1-7b30841ec5c5
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.RemoveFromContainers method (Visio)

Removes the shape from all lists and containers of which it is a member.


## Syntax

_expression_. `RemoveFromContainers`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

 **Nothing**


## Remarks

When you call the  **RemoveFromContainers** method, Microsoft Visio uses the **[ContainerProperties.ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property setting for each container to determine how to resize the container.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]