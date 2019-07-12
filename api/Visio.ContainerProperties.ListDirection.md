---
title: ContainerProperties.ListDirection property (Visio)
keywords: vis_sdr.chm17662600
f1_keywords:
- vis_sdr.chm17662600
ms.prod: visio
api_name:
- Visio.ContainerProperties.ListDirection
ms.assetid: 0024e464-a865-dfd2-9936-569827e529c0
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.ListDirection property (Visio)

Determines the primary list direction of the container shapes. Read/write.


## Syntax

_expression_.**ListDirection**

_expression_ An expression that returns a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Return value

**[VisListDirection](Visio.VisListDirection.md)**


## Remarks

The value of the **ListDirection** property can be one of the following **VisListDirection** constants.

|Constant|Value|Description|
|:-----|:-----|:-----|
| **visListDirLeftToRight**|0|Shapes are arranged horizontally, from left to right.|
| **visListDirRightToLeft**|1|Shapes are arranged horizontally, from right to left.|
| **visListDirTopToBottom**|2|Shapes are arranged vertically, from top to bottom.|
| **visListDirBottomToTop**|3|Shapes are arranged vertically, from bottom to top.|

If the container is not a list, Microsoft Visio returns an **Invalid Source** error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]