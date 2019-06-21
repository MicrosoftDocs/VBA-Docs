---
title: Viewer.ReviewerMarkupVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ReviewerMarkupVisible
ms.assetid: 3c365da2-1eac-0462-607b-be9923f62942
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ReviewerMarkupVisible property (Visio Viewer)

Gets or sets a value that indicates whether markup of the specified reviewer is visible in the drawing that is open in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**ReviewerMarkupVisible** (_ReviewerIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ReviewerIndex_|Required| **Long**|The index of the reviewer in the collection of reviewers.|

## Return value

**Boolean**


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1. The default is for reviewer markup to be visible for any reviewers in the drawing. If there are no reviewers in the drawing, or if you pass the index of a nonexistent reviewer, Visio Viewer returns an error.


## Example

The following code shows how to display the markup of the first reviewer in the collection of reviewers in the drawing that is open in Visio Viewer.

```vb
vsoViewer.ReviewerMarkupVisible(1)  = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]