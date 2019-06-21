---
title: Viewer.Zoom property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.Zoom
ms.assetid: 52bb7493-836e-1e1b-a91e-cb077f881c00
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.Zoom property (Visio Viewer)

Gets or sets the percentage of zoom for Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**Zoom**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Double**


## Remarks

Possible values for the **Zoom** property range from 1% through 400%, and also include Page, Width, and Last.


## Example

The following code gets the percentage of zoom in the drawing that is open in Visio Viewer.

```vb
Debug.Print "Zoom = "; vsoViewer.Zoom
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]