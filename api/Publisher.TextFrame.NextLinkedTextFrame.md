---
title: TextFrame.NextLinkedTextFrame property (Publisher)
keywords: vbapb10.chm3866648
f1_keywords:
- vbapb10.chm3866648
ms.prod: publisher
api_name:
- Publisher.TextFrame.NextLinkedTextFrame
ms.assetid: 5ba08ab5-8515-4efe-59a3-79a11f6a7c4e
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.NextLinkedTextFrame property (Publisher)

Returns or sets a **TextFrame** object representing the text frame to which text flows from the specified text frame. Read/write.


## Syntax

_expression_.**NextLinkedTextFrame**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Return value

TextFrame


## Remarks

If the specified text frame is not part of a chain of linked frames or is the last in a chain of linked frames, this property returns nothing.


## Example

The following example returns the next linked text frame of shape three on page one of the active publication and sets its font to Times New Roman.

```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.NextLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]