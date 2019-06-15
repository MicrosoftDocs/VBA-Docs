---
title: TextRange.Duplicate property (Publisher)
keywords: vbapb10.chm5308466
f1_keywords:
- vbapb10.chm5308466
ms.prod: publisher
api_name:
- Publisher.TextRange.Duplicate
ms.assetid: 545dbfdb-4cd5-99b1-1ba3-b723e8d7b827
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Duplicate property (Publisher)

Returns a **TextRange** object that represents a duplicate of the specified text range.


## Syntax

_expression_.**Duplicate**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

TextRange


## Example

This example sets the value of a string variable to the contents of the specified text box on the first page of the active publication. It then creates a new page with a text box and sets the contents of the new text box equal to the value of the string variable.

```vb
Sub DuplicateTextBoxContents() 
 Dim strDuplicate As String 
 Dim pagNew As Page 
 
 With ThisDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 strDuplicate = .Duplicate 
 End With 
 
 Set pagNew = ThisDocument.Pages.Add(Count:=1, After:=1) 
 
 pagNew.Shapes.AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=200, Height:=200).TextFrame _ 
 .TextRange.Text = strDuplicate 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]