---
title: TextRange.Story property (Publisher)
keywords: vbapb10.chm5308470
f1_keywords:
- vbapb10.chm5308470
ms.prod: publisher
api_name:
- Publisher.TextRange.Story
ms.assetid: 833f9537-5c11-a4d5-907a-777eaecb89d2
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Story property (Publisher)

Returns a **[Story](publisher.story.md)** object that represents the story properties in a text range.


## Syntax

_expression_.**Story**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Example

This example returns the story in the selected text range, and if it is in a text frame, inserts text into the text range.

```vb
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then 
 .TextRange.InsertAfter NewText:=vbLf & "This is a test." 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]