---
title: VisWebPageSettings.PanAndZoom property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.PanAndZoom
ms.assetid: 83d1ac9d-e489-0656-a573-ebadd6e06156
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.PanAndZoom property

Determines whether the **Pan and Zoom** control for zooming in and out of the page is displayed on a webpage. Read/write.


## Syntax

_expression_.**PanAndZoom**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

**PanAndZoom** returns non-zero (**True**) if the **Pan and Zoom** control is displayed after the drawing is exported to a webpage; otherwise, it returns zero (**False**). The default is **True**.

Set **PanAndZoom** to a non-zero value (**True**) to display the **Pan and Zoom** control after the drawing is exported to a webpage; otherwise, set it to zero (**False**).

The **PanAndZoom** property corresponds to the **Pan and Zoom** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

> [!NOTE] 
> The **Pan and Zoom** control is supported for the VML output format in Microsoft Internet Explorer 5 and later. The **Pan and Zoom** control is not available in SVG, JPG, GIF, and PNG output formats.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]