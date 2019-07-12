---
title: Document.Container property (Visio)
keywords: vis_sdr.chm10513295
f1_keywords:
- vis_sdr.chm10513295
ms.prod: visio
api_name:
- Visio.Document.Container
ms.assetid: a5b2c90e-f9e0-cc09-8388-566729c1c4eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Container property (Visio)

Returns an **IDispatch** interface on the ActiveX container in which the document is contained or **Nothing** if the document is not in a container. Read-only.


## Syntax

_expression_.**Container**

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Return value

Object


## Remarks

The interface returned is the result of querying the **IOleContainer** interface provided by the containing object for **IDispatch**.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]