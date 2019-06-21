---
title: Viewer.LayerCount property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.LayerCount
ms.assetid: 83871b37-9c5b-9da2-8869-61a2284ae1c0
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.LayerCount property (Visio Viewer)

Gets the number of layers in the current drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**LayerCount**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Long**


## Remarks

If there are no layers in the drawing, the **LayerCount** property returns 0.


## Example

The following code gets the count of layers in the drawing open in Visio Viewer.

```vb
Debug.Print vsoViewer.LayerCount
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]