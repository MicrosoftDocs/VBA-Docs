---
title: ContainerProperties.ContainerType property (Visio)
keywords: vis_sdr.chm17662590
f1_keywords:
- vis_sdr.chm17662590
ms.prod: visio
api_name:
- Visio.ContainerProperties.ContainerType
ms.assetid: ba3ead35-a6da-5978-e852-4362e5ca230e
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.ContainerType property (Visio)

Determines whether the container is a normal container or a list container. Read-only.


## Syntax

_expression_.**ContainerType**

 _expression_ An expression that returns a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Return value

 **[VisContainerTypes](Visio.VisContainerTypes.md)**


## Remarks

The value of the  **ContainerType** property can be one of the following **VisContainerTypes** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visContainerTypeNormal**|0|Member shapes are not arranged in a list.|
| **visContainerTypeList**|1|Member shapes are arranged in a list.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]