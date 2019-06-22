---
title: Viewer.ReviewerCount property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ReviewerCount
ms.assetid: 5ab6cae5-ea59-bb72-1fb2-04aebc5ae5cc
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ReviewerCount property (Visio Viewer)

Gets the count of reviewers in the current document open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**ReviewerCount**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Long**


## Remarks

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the number of reviewers in the drawing open in Visio Viewer and displays it in the Immediate window.

```vb
Debug.Print vsoViewer.ReviewerCount
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]