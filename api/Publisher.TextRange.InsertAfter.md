---
title: TextRange.InsertAfter method (Publisher)
keywords: vbapb10.chm5308448
f1_keywords:
- vbapb10.chm5308448
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertAfter
ms.assetid: f647be29-68c7-b221-adf1-fa233583e74e
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.InsertAfter method (Publisher)

Returns a **TextRange** object that represents text appended to the end of a text range.


## Syntax

_expression_.**InsertAfter** (_NewText_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_NewText_|Required| **String**|The text to be inserted.|

## Return value

TextRange


## Example

This example adds the Microsoft Publisher build number to the end of the first shape on the first page of the active publication. This example assumes that the specified shape is a text frame and not another type of shape.

```vb
Sub AppendText() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .TextFrame.TextRange.InsertAfter _ 
 NewText:="Microsoft Publisher Build : " & Build 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]