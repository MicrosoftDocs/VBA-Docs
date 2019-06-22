---
title: Viewer.ReviewerID property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ReviewerID
ms.assetid: dc6c8175-9cfb-5f31-8572-d7ead88d96a4
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ReviewerID property (Visio Viewer)

Gets the ID of the specified reviewer in the drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**ReviewerID** (_ReviewerIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ReviewerIndex_|Required| **Long**|The index of the reviewer in the collection of reviewers.|

## Return value

**Long**


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the ID of the reviewer at index position 1 in the drawing open in Visio Viewer.

```vb
Debug.Print vsoViewer.ReviewerID(1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]