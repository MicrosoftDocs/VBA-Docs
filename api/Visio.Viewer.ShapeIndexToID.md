---
title: Viewer.ShapeIndexToID property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ShapeIndexToID
ms.assetid: 9f43bcb1-1c10-3759-e740-bc4ae04a51be
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ShapeIndexToID property (Visio Viewer)

Gets the ID of the shape at the specified index position in the collection of shapes in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**ShapeIndexToID** (_ShapeIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index position of the shape in the collection of shapes.|

## Return value

**Long**


## Remarks

The collection of shapes in Visio Viewer is one-based, so the first shape in the collection is at index position 1.


## Example

The following code gets the ID of all the shapes in the drawing that is open in Visio Viewer.

```vb
Dim intCounter As Integer

    For intCounter = 1 To Viewer1.ShapeCount

    Debug.Print Viewer1.ShapeIndexToID(intCounter)

Next intCounter

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]