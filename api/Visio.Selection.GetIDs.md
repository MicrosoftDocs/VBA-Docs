---
title: Selection.GetIDs method (Visio)
keywords: vis_sdr.chm11160200
f1_keywords:
- vis_sdr.chm11160200
ms.prod: visio
api_name:
- Visio.Selection.GetIDs
ms.assetid: 79b1fb3f-eb53-2640-a988-6e79b067f228
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.GetIDs method (Visio)

Gets the shape IDs of the shapes in the selection.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**GetIDs**(_ShapeIDs()_)

_expression_ An expression that returns a **[Selection](Visio.Selection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShapeIDs()_|Required| **Long**|Out parameter. An array of shape IDs of type **Long** corresponding to the shapes in the selection.|

## Return value

Nothing


## Remarks

Microsoft Visio uses ID numbers to identify shapes, recordsets, and data rows. Shape IDs are unique only within the scope of the page they are on. After you determine these shape IDs, you can pass them to the **[Page.LinkShapesToDataRows](Visio.Page.LinkShapesToDataRows.md)** method to specify exactly how the shapes in your diagram should link to data rows in the available data recordsets. Shape IDs are unique within the scope of a particular page.

To determine the shape ID for a shape that is part of a selection, use the **Selection.GetIDs** method.

The set of shape IDs returned is determined by the setting of the **[Selection.IterationMode](Visio.Selection.IterationMode.md)** property.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetIDs** method to get the IDs of shapes in a selection and print the IDs in the Immediate window. It selects all the shapes in the active window.


```vb
Public Sub GetIDs_Example() 
 
    Dim vsoSelection As Visio.Selection 
    Dim lngShapeID As Long
    Dim intShapeIndex As Integer 
     
    ActiveWindow.DeselectAll 
    ActiveWindow.SelectAll 
     
    Set vsoSelection = ActiveWindow.Selection 
     
    Call vsoSelection.GetIDs(lngShapeIDs) 
     
    For intShapeIndex = LBound(lngShapeIDs) To UBound(lngShapeIDs) 
        lngShapeID = lngShapeIDs(intShapeIndex)
        Debug.Print lngShapeID 
    Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]