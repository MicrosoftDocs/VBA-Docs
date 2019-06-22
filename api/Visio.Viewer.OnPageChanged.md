---
title: Viewer.OnPageChanged event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnPageChanged
ms.assetid: de64b0e2-372c-f1c4-15c6-d6c3a4da8368
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnPageChanged event (Visio Viewer)

Occurs when the active page is changed in Microsoft Visio Viewer.


## Syntax

_expression_.**OnPageChanged** (_PageIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PageIndex_|Required| **Long**|The index of the new page.|

## Return value

Nothing


## Remarks

The collection of pages in the Viewer is one-based, so the index of the first page in the collection is 1. 

You can change the page programmatically in Visio Viewer by setting the value of the **[CurrentPageIndex](Visio.Viewer.CurrentPageIndex.md)** property.


## Example

The following code shows how to use the **OnPageChanged** event to print a message in the Immediate window stating that the page has changed and identifying the new page.

```vb
Private Sub vsoViewer_OnPageChanged(ByVal PageIndex As Long)

    Debug.Print "Page changed to"; vsoViewer.CurrentPageIndex

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]