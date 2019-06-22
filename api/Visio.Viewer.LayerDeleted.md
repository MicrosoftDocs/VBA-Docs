---
title: Viewer.LayerDeleted property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.LayerDeleted
ms.assetid: cb7ea0ab-fdf8-2621-5ebc-edab2d9869f8
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.LayerDeleted property (Visio Viewer)

Gets a value that indicates whether the layer at the specified index in the drawing open in Microsoft Visio Viewer is deleted. Read-only.


## Syntax

_expression_.**LayerDeleted** (_LayerIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_LayerIndex_|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

## Return value

**Boolean**


## Remarks

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, the **LayerDeleted** property returns **False**.


## Example

The following code gets a value that indicates whether the layer at index position 1 is deleted.

```vb
Debug.Print vsoViewer.LayerDeleted(1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]