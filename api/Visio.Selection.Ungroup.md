---
title: Selection.Ungroup Method (Visio)
keywords: vis_sdr.chm11116625
f1_keywords:
- vis_sdr.chm11116625
ms.prod: visio
api_name:
- Visio.Selection.Ungroup
ms.assetid: b9f14342-e885-1399-83ed-59189f5cbec3
ms.date: 06/08/2017
---


# Selection.Ungroup Method (Visio)

Ungroups a group.


## Syntax

 _expression_. `Ungroup`

 _expression_ A variable that represents a [Selection](./Visio.Selection.md) object.


### Return value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Ungroup** method.


```vb
 
Sub Ungroup_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 Dim vsoGroup As Visio.Shape 
 Dim vsoSelection As Visio.Selection 
 
 'Draw two rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 2, 2, 1) 
 Set vsoShape2 = ActivePage.DrawRectangle(1, 4, 2, 3) 
 
 'Add a copy of one of the rectangles to the page. 
 ActivePage.Drop vsoShape1, 3.5, 3.5 
 Set vsoShape3 = ActivePage.Shapes(3) 
 
 'Deselect all shapes, and then select all the shapes on the page. 
 Set vsoSelection = ActiveWindow.Selection 
 vsoSelection.Select vsoShape1, visDeselectAll + visSelect 
 vsoSelection.Select vsoShape2, visSelect 
 vsoSelection.Select vsoShape3, visSelect 
 
 'Group all the shapes into a group shape. 
 Set vsoGroup = vsoSelection.Group 
 
 'Ungroup the shapes. 
 vsoGroup.Ungroup 
 
End Sub
```


