---
title: Story.TextRange property (Publisher)
keywords: vbapb10.chm5832712
f1_keywords:
- vbapb10.chm5832712
ms.prod: publisher
api_name:
- Publisher.Story.TextRange
ms.assetid: c948da79-ea67-0c8c-1df3-2b32499ea9b3
ms.date: 06/14/2019
localization_priority: Normal
---


# Story.TextRange property (Publisher)

Returns a **[TextRange](Publisher.TextRange.md)** object that represents the text that is attached to a shape and properties and methods for manipulating the text.


## Syntax

_expression_.**TextRange**

_expression_ A variable that represents a **[Story](Publisher.Story.md)** object.


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