---
title: Selection Object (Publisher)
keywords: vbapb10.chm917503
f1_keywords:
- vbapb10.chm917503
ms.prod: publisher
api_name:
- Publisher.Selection
ms.assetid: 1ebee88b-a39e-ea3a-48b0-6205621853af
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection Object (Publisher)

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the publication, or it represents the cursor if nothing in the publication is selected. There can only be one  **Selection** object per publication window pane, and only one **Selection** object in the entire application can be active.
 


## Example

Use the  **[Selection](Publisher.Document.Selection.md)** property to return the **Selection** object. If no object qualifier is used with the **Selection** property, Microsoft Publisher returns the selection from the active pane of the active publication window. The following example copies the current selection from the active publication.
 

 

```vb
Sub CopySelection() 
 Selection.ShapeRange.Copy 
End Sub
```

The following example determines what type of item is selected and, if it is an autoshape, fills the first shape in the selection with color. This example assumes there is at least one item selected in the active pubication.
 

 



```vb
Sub SelectedShape() 
 If Selection.Type = pbSelectionShape Then 
 Selection.ShapeRange.Item(1).Fill.ForeColor _ 
 .RGB = RGB(Red:=200, Green:=20, Blue:=255) 
 End If 
End Sub
```

The following example copies the selection and pastes it into the first shape on the second page of the active publication.
 

 



```vb
Sub CopyPasteSelection() 
 Selection.TextRange.Copy 
 With ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
 .Collapse Direction:=pbCollapseEnd 
 .InsertAfter NewText:=vbLf 
 .Paste 
 End With 
End Sub
```


## Methods



|Name|
|:-----|
|[Unselect](Publisher.Selection.Unselect.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.Selection.Application.md)|
|[ChildShapeRange](Publisher.Selection.ChildShapeRange.md)|
|[Parent](Publisher.Selection.Parent.md)|
|[ShapeRange](Publisher.Selection.ShapeRange.md)|
|[TableCellRange](Publisher.Selection.TableCellRange.md)|
|[TextRange](Publisher.Selection.TextRange.md)|
|[Type](Publisher.Selection.Type.md)|

