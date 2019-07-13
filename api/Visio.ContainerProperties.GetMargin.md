---
title: ContainerProperties.GetMargin method (Visio)
keywords: vis_sdr.chm17662300
f1_keywords:
- vis_sdr.chm17662300
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetMargin
ms.assetid: c0e224a1-f7a6-e16c-a99c-766a5a4ac207
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.GetMargin method (Visio)

Returns the minimal distance, in the specified units, between the edges of the container or list and those of its member shapes.


## Syntax

_expression_.**GetMargin** (_MarginUnits_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarginUnits_|Required| **[VisUnitCodes](Visio.visunitcodes.md)**|The units in which to measure the margin.|

## Return value

**Double**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]