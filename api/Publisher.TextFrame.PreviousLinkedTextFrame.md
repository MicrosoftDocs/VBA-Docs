---
title: TextFrame.PreviousLinkedTextFrame property (Publisher)
keywords: vbapb10.chm3866656
f1_keywords:
- vbapb10.chm3866656
ms.prod: publisher
api_name:
- Publisher.TextFrame.PreviousLinkedTextFrame
ms.assetid: 00947ec3-fcff-4451-491b-5b7748ccb74e
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.PreviousLinkedTextFrame property (Publisher)

Returns a **TextFrame** object representing the text frame from which text flows to the specified text frame.


## Syntax

_expression_.**PreviousLinkedTextFrame**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Return value

TextFrame


## Remarks

If the specified text frame is not part of a chain of linked frames or is the first in a chain of linked frames, this property returns nothing.


## Example

The following example returns the previously linked text frame of shape three on page one of the active publication and sets its font to Times New Roman.

```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.PreviousLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]