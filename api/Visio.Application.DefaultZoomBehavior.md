---
title: Application.DefaultZoomBehavior property (Visio)
keywords: vis_sdr.chm10052070
f1_keywords:
- vis_sdr.chm10052070
ms.prod: visio
api_name:
- Visio.Application.DefaultZoomBehavior
ms.assetid: 59f26e36-90e3-defa-be04-b7a8ce710eeb
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.DefaultZoomBehavior property (Visio)

Determines the zoom behavior for all new Microsoft Visio documents and drawing windows. Read/write.


## Syntax

_expression_.**DefaultZoomBehavior**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

VisZoomBehavior


## Remarks

To set zoom behavior for an existing document, or for a particular window, use the **ZoomBehavior** property of the document and window, respectively.

The constants declared by the Visio type library in **[VisZoomBehavior](visio.viszoombehavior.md)** are valid values for the **DefaultZoomBehavior** property.

> [!NOTE] 
> The default behavior (**visZoomInPlaceContainer**) is different from the behavior used in Microsoft Visio 2002, but is the same as that in earlier versions of Visio. To replicate the behavior seen in Microsoft Visio 2002, set this value to **visZoomVisio**.

If this value is set to the default, **visZoomInPlaceContainer**, Visio uses the container's **IOleCommandTarget** interface to perform the zoom and forces a fit-to-window zoom within the in-place window. If the container does not support **IOleCommandTarget**, no zooming occurs.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]