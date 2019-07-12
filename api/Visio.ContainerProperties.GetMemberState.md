---
title: ContainerProperties.GetMemberState method (Visio)
keywords: vis_sdr.chm17662330
f1_keywords:
- vis_sdr.chm17662330
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetMemberState
ms.assetid: 04103f79-7f28-7584-3bab-0c1d140f6b52
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.GetMemberState method (Visio)

Returns the membership state of the specified shape with respect to the container shape.


## Syntax

_expression_.**GetMemberState** (_Shape_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[Shape](Visio.Shape.md)**|The shape for which to get the membership state.|

## Return value

[VisContainerMemberState](Visio.VisContainerMemberState.md)


## Remarks

**GetMemberState** can return one of the following **VisContainerMemberState** constants.

|Constant|Value|Description|
|:-----|:-----|:-----|
| **visContainerMemberNotAMember**|0|The shape is not a member of the container.|
| **visContainerMemberInterior**|1|The member shape is within the bounds of the container.|
| **visContainerMemberOnBoundary**|2|The member shape is on the boundary of the container.|
| **visContainerMemberOutside**|3|The member shape is outside the bounds of the container.|
| **visContainerMemberInList**|4|The member shape is a list member.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]