---
title: Shape.AddToContainers method (Visio)
keywords: vis_sdr.chm11262215
f1_keywords:
- vis_sdr.chm11262215
ms.prod: visio
api_name:
- Visio.Shape.AddToContainers
ms.assetid: ddd3f532-cbbf-3c63-0e02-49f4ea8ca90c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.AddToContainers method (Visio)

Adds the shape to all underlying containers that allow it as a member.


## Syntax

_expression_. `AddToContainers`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

 **Nothing**


## Remarks

When you call the  **AddToContainers** method, Microsoft Visio uses the setting of the **[ContainerProperties.ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property for each container to determine how the container resizes.

If the underlying container is a list, the shape is added as a normal container member, and not as a list member.

The  **AddToContainers** method works only if the shape sits at least partially on top of a container that does not already contain it.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]