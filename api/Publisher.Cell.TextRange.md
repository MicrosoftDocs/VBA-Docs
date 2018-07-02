---
title: Cell.TextRange Property (Publisher)
keywords: vbapb10.chm5111833
f1_keywords:
- vbapb10.chm5111833
ms.prod: publisher
api_name:
- Publisher.Cell.TextRange
ms.assetid: 31aa92d1-852f-3742-defa-94485411bcc3
ms.date: 06/08/2017
---


# Cell.TextRange Property (Publisher)

Returns a  **[TextRange](Publisher.TextRange.md)** object that represents the text that is attached to a shape and properties and methods for manipulating the text.


## Syntax

 _expression_. **TextRange**

 _expression_ A variable that represents a  **Cell** object.


## Example

The following example adds text to the text frame of shape one in the active publication, and then formats the new text. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```

The following example adds a rectangle to the active publication and adds text to it.




```vb
Sub AddTextToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
 Left:=72, Top:=72, Width:=250, Height:=140) 
 .TextFrame.TextRange.Text = "Here is some test text" 
 End With 
End Sub
```


