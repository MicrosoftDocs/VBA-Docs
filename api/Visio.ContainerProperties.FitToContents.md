---
title: ContainerProperties.FitToContents method (Visio)
keywords: vis_sdr.chm17662295
f1_keywords:
- vis_sdr.chm17662295
ms.prod: visio
api_name:
- Visio.ContainerProperties.FitToContents
ms.assetid: 09169624-f1fd-66a3-0be2-738d808d540a
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.FitToContents method (Visio)

Forces the container to resize so as to tightly include all member shapes, including any applicable margins between the container and the shapes.


## Syntax

_expression_.**FitToContents**

_expression_ A variable that represents a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Return value

 **Nothing**


## Remarks

Calling the **FitToContents** method has no effect on the **[ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property setting for the current session of Microsoft Visio.

The **FitToContents** method does not work for list shapes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]