---
title: Viewer.ShapeIDToIndex property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ShapeIDToIndex
ms.assetid: ffb4020a-cc45-f012-ee21-abd9805bf99f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ShapeIDToIndex property (Visio Viewer)

Gets the index in the collection of shapes of the shape with the specified ID in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**ShapeIDToIndex** (_ShapeID_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeID_|Required| **Long**|The ID of the shape.|

## Return value

**Long**


## Remarks

The collection of shapes in Visio Viewer is one-based, so the first shape in the collection is at index position 1.


## Example

The following code gets the index position of all the shapes in the drawing that is open in Visio Viewer.

```vb
Dim intCounter As Integer

    For intCounter = 1 To Viewer1.ShapeCount

    Debug.Print Viewer1.ShapeIDToIndex(intCounter)

Next intCounter

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]