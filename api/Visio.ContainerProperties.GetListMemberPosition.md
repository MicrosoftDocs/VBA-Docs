---
title: ContainerProperties.GetListMemberPosition method (Visio)
keywords: vis_sdr.chm17662325
f1_keywords:
- vis_sdr.chm17662325
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetListMemberPosition
ms.assetid: 4fb6ab3b-b369-5e33-0b4f-50754d31f39d
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.GetListMemberPosition method (Visio)

Returns the ordinal position of the specified shape in the list.


## Syntax

_expression_.**GetListMemberPosition** (_ShapeMember_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapeMember_|Required| **[Shape](Visio.Shape.md)**|The list member shape for which you want to get the position in the container list.|

## Return value

**Long**


## Remarks

If the specified shape is not a member of the list, Microsoft Visio returns an Invalid Parameter error. 

List position is one-based.

If the container is not a list, Visio returns an Invalid Source error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]