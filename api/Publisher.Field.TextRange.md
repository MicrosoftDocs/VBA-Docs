---
title: Field.TextRange property (Publisher)
keywords: vbapb10.chm6094852
f1_keywords:
- vbapb10.chm6094852
ms.prod: publisher
api_name:
- Publisher.Field.TextRange
ms.assetid: 09279cc7-3911-3b8d-51f2-b26494220c68
ms.date: 06/07/2019
localization_priority: Normal
---


# Field.TextRange property (Publisher)

Returns a **[TextRange](Publisher.TextRange.md)** object that represents the text that is attached to a shape and properties and methods for manipulating the text.


## Syntax

_expression_.**TextRange**

_expression_ A variable that represents a **[Field](Publisher.Field.md)** object.


## Example

The following example adds text to the text frame of shape one in the active publication, and then formats the new text. This example assumes that there is at least one shape on the first page of the active publication.

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

<br/>

The following example adds a rectangle to the active publication and adds text to it.

```vb
Sub AddTextToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
 Left:=72, Top:=72, Width:=250, Height:=140) 
 .TextFrame.TextRange.Text = "Here is some test text" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]