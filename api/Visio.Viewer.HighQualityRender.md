---
title: Viewer.HighQualityRender property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.HighQualityRender
ms.assetid: 39f59bc2-36ad-7c74-97de-85a486eb42c3
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.HighQualityRender property (Visio Viewer)

Gets or sets a value that indicates whether high-quality rendering is enabled in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**HighQualityRender**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

High-quality rendering is slower but produces output that looks better.

The default is for high-quality rendering to be enabled (property value set to **True**).


## Example

The following code gets a value that indicates whether high-quality rendering is enabled in Visio Viewer.

```vb
Debug.Print vsoViewer.HighQualityRender
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]