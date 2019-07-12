---
title: ContainerProperties.RemoveMember method (Visio)
keywords: vis_sdr.chm17662335
f1_keywords:
- vis_sdr.chm17662335
ms.prod: visio
api_name:
- Visio.ContainerProperties.RemoveMember
ms.assetid: 953beb58-ea8a-7c1f-20c1-0fe4de23e831
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.RemoveMember method (Visio)

Removes a shape or set of shapes from the container.


## Syntax

_expression_.**RemoveMember** (_ObjectToRemove_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectToRemove_|Required| **[UNKNOWN]**|The shape or shapes to remove from the container. Can be a **[Shape](Visio.Shape.md)** or **[Selection](Visio.Selection.md)** selection.|

## Return value

**Nothing**


## Remarks

The **RemoveMember** method removes from the container the shapes specified in the _ObjectToRemove_ parameter.

If the container is a list, Microsoft Visio removes the shapes specified in  _ObjectToRemove_ both from the list (if it is a list member) and from the list container.

If the **[ContainerProperties.LockMembership](Visio.ContainerProperties.LockMembership.md)** property is **True**, Visio returns a Disabled error.

If  _ObjectToRemove_ does not contain top-level shapes on the page, Visio returns an Invalid Parameter error. However, if _ObjectToRemove_ is not a container member, Visio does not return an error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]