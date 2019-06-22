---
title: Viewer.LayerColorOverride property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.LayerColorOverride
ms.assetid: 378cd05b-50b0-2169-9419-0d489860f0ad
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.LayerColorOverride property (Visio Viewer)

Gets or sets a value that indicates whether the color of the specified layer is reset to the default color in the current drawing in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**LayerColorOverride** (_LayerIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_LayerIndex_|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

## Return value

**Boolean**


## Remarks

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, or if you pass the index of a nonexistent layer, the **LayerColorOverride** property returns **False**. The default value is **True**.


## Example

The following code shows how to override the color of the layer at index position 1.

```vb
vsoViewer.LayerColorOverride(1) = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]