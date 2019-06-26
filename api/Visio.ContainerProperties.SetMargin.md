---
title: ContainerProperties.SetMargin method (Visio)
keywords: vis_sdr.chm17662305
f1_keywords:
- vis_sdr.chm17662305
ms.prod: visio
api_name:
- Visio.ContainerProperties.SetMargin
ms.assetid: 008dbfe9-53d9-17a6-c441-b30d5a691716
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.SetMargin method (Visio)

Sets the gap between the container and member shapes to the specified size, in the specified units.


## Syntax

_expression_.**SetMargin** (_MarginUnits_, _MarginSize_)

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarginUnits_|Required| **[VisUnitCodes](Visio.visunitcodes.md)**|The units in which to measure the margin.|
| _MarginSize_|Required| **Double**|The size of the margin.|

## Return value

 **Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]