---
title: Viewer.OnReviewerChanged event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnReviewerChanged
ms.assetid: a705878b-cb2e-5b5c-01ae-e0fca790c0d5
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnReviewerChanged event (Visio Viewer)

Occurs when the visibility of a particular reviewer's markup (comments) is changed in Microsoft Visio Viewer.


## Syntax

_expression_.**OnReviewerChanged** (_ReviewerIndex_, _ReviewerVisible_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ReviewerIndex_|Required| **Long**|The index of the reviewer.|
|_ReviewerVisible_|Required| **Boolean**|Indicates whether markup of the reviewer is visible in the user interface.|

## Return value

Nothing


## Remarks

The collection of reviewers in the Viewer is one-based, so the index of the first reviewer in the collection is 1. 

You can specify whether markup of a reviewer is visible in the Visio Viewer user interface by setting the **[ReviewerMarkupVisible](Visio.Viewer.ReviewerMarkupVisible.md)** property.

The **OnReviewerChanged** event occurs when markup overlays of a specific reviewer are set to be visible or not visible. 

The **[OnMarkupOverlaysVisibleChanged](Visio.Viewer.OnMarkupOverlaysVisibleChanged.md)** event occurs when markup overlays for all reviewers in a drawing are set to be visible or not visible.


## Example

The following code shows how to use the **OnReviewerChanged** event to print a message in the Immediate window identifying the reviewer and stating the visibility status of the reviewer's markup.

```vb
Private Sub vsoViewer_OnReviewerChanged(ByVal ReviewerIndex As Long, ByVal ReviewerVisible As Boolean)

    If ReviewerVisible Then

        Debug.Print "Reviewer "; vsoViewer.ReviewerName(ReviewerIndex); " markup is visible."

    Else

        Debug.Print "Reviewer "; vsoViewer.ReviewerName(ReviewerIndex); " markup is not visible."

    End If

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]