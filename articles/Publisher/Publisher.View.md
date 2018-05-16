---
title: View Object (Publisher)
keywords: vbapb10.chm393215
f1_keywords:
- vbapb10.chm393215
ms.prod: publisher
api_name:
- Publisher.View
ms.assetid: a865cf48-cd68-5789-6855-c09c05b7634b
ms.date: 06/08/2017
---


# View Object (Publisher)

Contains the view attributes (show all, field shading, table gridlines, and so on) for a window or pane.
 


## Example

Use the  **[ActiveView](Publisher.Document.ActiveView.md)** property to return the **View** object. The following example specifies the zoom setting.
 

 

```
Sub ZoomFitSelection() 
 ActiveDocument.ActiveView.Zoom = pbZoomFitSelection 
End Sub
```

The following examples zoom in and out, respectively, on the active view.
 

 



```
Sub ViewZoomIn() 
 ActiveDocument.ActiveView.ZoomIn 
End Sub 
 
Sub ViewZoomOut() 
 ActiveDocument.ActiveView.ZoomOut 
End Sub
```

The following example scrolls the active view to the specified shape.
 

 



```
Sub ScrollToShape() 
 Dim shpOne As Shape 
 
 Set shpOne = ActiveDocument.Pages(1).Shapes(1) 
 ActiveDocument.ActiveView.ScrollShapeIntoView Shape:=shpOne 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ScrollShapeIntoView](Publisher.View.ScrollShapeIntoView.md)|
|[ZoomIn](Publisher.View.ZoomIn.md)|
|[ZoomOut](Publisher.View.ZoomOut.md)|

## Properties



|**Name**|
|:-----|
|[ActivePage](Publisher.View.ActivePage.md)|
|[Application](Publisher.View.Application.md)|
|[Parent](Publisher.View.Parent.md)|
|[Zoom](Publisher.View.Zoom.md)|

