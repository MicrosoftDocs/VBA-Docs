---
title: Viewer.OnSelectionChanged event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnSelectionChanged
ms.assetid: 825a9f43-8a7f-7237-af84-3f13b8d19a04
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnSelectionChanged event (Visio Viewer)

Occurs when the shape selection is changed in Microsoft Visio Viewer.


## Syntax

_expression_.**OnSelectionChanged** (_ShapeIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index of the newly selected shape.|

## Return value

Nothing


## Remarks

The collection of shapes in the Viewer is one-based, so the first shape in the collection has an index of 1.

You can change the selected shape in Visio Viewer programmatically by using the **[SelectShape](Visio.Viewer.SelectShape.md)** method.


## Example

The following code shows how to use the **OnSelectionChanged** event to print the ID of the newly selected shape in Visio Viewer in the Immediate window.

```vb
Private Sub vsoViewer_OnSelectionChanged(ByVal ShapeIndex As Long)

    Debug.Print "Selected shape changed to shape " & vsoViewer.SelectedShapeIndex & "."

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]