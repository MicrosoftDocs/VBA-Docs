---
title: VisReplaceFlags enumeration (Visio)
ms.prod: visio
ms.assetid: cf270178-f939-7eb4-b8e1-3b4153aff221
ms.date: 06/08/2017
localization_priority: Normal
---


# VisReplaceFlags enumeration (Visio)

Shape-replacement flags sent to the [Shape.ReplaceShape](Visio.shape.replaceshape.md) and [Selection.ReplaceShape](Visio.selection.replaceshape.md) methods and returned by the [ReplaceShapesEvent.ReplaceFlags](Visio.replaceshapesevent.replaceflags.md) property, singly or in combination.


## Members



|Name|Value|Description|
|:-----|:-----|:-----|
||||
| **visReplaceShapeDefault**| **0**|Use the behavior specified by the ShapeSheet cells ReplaceLockText, ReplaceLockShapeData, and ReplaceLockFormat, all in the Change Shape Behavior section.|
| **visReplaceShapeKeepBasic**| **1**|Override the behavior specified by the following ShapeSheet cells, all in the Change Shape Behavior section: behave as if ReplaceLockText = 0, ReplaceLockShapeData = 0, and ReplaceLockFormat = 0.|
| **visReplaceShapeLockFormat**| **8**|Override the behavior specified by the ReplaceLockFormat cell in the Change Shape Behavior section: behave as if ReplaceLockFormat = 1.|
| **visReplaceShapeLockShapeData**| **4**|Override the behavior specified by the ReplaceLockShapeData cell in the Change Shape Behavior section: behave as if ReplaceLockShapeData = 1.|
| **visReplaceShapeLockText**| **2**|Override the behavior specified by the ReplaceLockText cell in the Change Shape Behavior section: behave as if ReplaceLockText = 1.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]