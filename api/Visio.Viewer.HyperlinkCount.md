---
title: Viewer.HyperlinkCount property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.HyperlinkCount
ms.assetid: 06c06812-25a6-779d-3af4-821538493c4f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.HyperlinkCount property (Visio Viewer)

Gets the count of hyperlinks associated with the shape at the specified index in the document open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**HyperlinkCount** (_ShapeIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index of the specified shape in the collection of shapes in the drawing open in Visio Viewer.|

## Return value

**Long**


## Remarks

The collection of shapes is one-based, so the index of the first shape in the collection is 1.


## Example

The following code gets the count of hyperlinks associated with the first shape in the collection of shapes in the drawing open in Visio Viewer.

```vb
Debug.Print vsoViewer.HyperlinkCount(1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]