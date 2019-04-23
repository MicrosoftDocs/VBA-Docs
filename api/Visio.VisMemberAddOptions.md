---
title: VisMemberAddOptions enumeration (Visio)
keywords: vis_sdr.chm70615
f1_keywords:
- vis_sdr.chm70615
ms.prod: visio
api_name:
- Visio.VisMemberAddOptions
ms.assetid: e6833a87-2d08-19a4-c2f9-86803ca4e4ba
ms.date: 06/08/2017
localization_priority: Normal
---


# VisMemberAddOptions enumeration (Visio)

Specifies whether to expand the container to accommodate the new member(s) or to resize it automatically according to the default settings; constants passed to the  **[ContainerProperties.AddMember](Visio.ContainerProperties.AddMember.md)** method.



|Name|Value|Description|
|:-----|:-----|:-----|
| **visMemberAddUseResizeSetting**|0|Defer to the setting of the  **[ContainerProperties.ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property.|
| **visMemberAddExpandContainer**|1|Expand the container to fit the incoming shape(s).|
| **visMemberAddDoNotExpand**|2|Do not expand the container to fit the incoming shape(s).|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]