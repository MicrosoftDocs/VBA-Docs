---
title: Master.VisualBoundingBox method (Visio)
ms.assetid: 478d636f-e741-cf6b-3e16-b5faf70a9f14
ms.date: 06/08/2017
ms.prod: visio
localization_priority: Normal
---


# Master.VisualBoundingBox method (Visio)

Returns the bounding rectangle of the virtual container that has all the shapes of the given master. Introduced in Office 2016.


## Syntax

_expression_.**VisualBoundingBox** (_Flags_, _Flags_, _lpr8Left_, _lpr8Bottom_, _lpr8Right_, _lpr8Top_)

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters

|Name|Optional/Requires|Data Type|Description|
|:-----|:-----|:-----|:-----|
| _Flags_|Required|INT16|A **[VisBoundingBoxArgs](Visio.visboundingboxargs.md)** constant that describes the returned rectangle.|
| _lpr8Left_|Required|DOUBLE|Left position values for the virtual bounding box.|
| _lpr8Bottom_|Required|DOUBLE|Bottom position values for the virtual bounding box.|
| _lpr8Right_|Required|DOUBLE|Right position values for the virtual bounding box.|
| _lpr8Top_|Required|DOUBLE|Top position values for the virtual bounding box.|

## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]