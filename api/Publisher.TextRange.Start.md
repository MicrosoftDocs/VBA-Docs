---
title: TextRange.Start property (Publisher)
keywords: vbapb10.chm5308433
f1_keywords:
- vbapb10.chm5308433
ms.prod: publisher
api_name:
- Publisher.TextRange.Start
ms.assetid: 40604058-7c3e-b4c7-c793-bbf09091b4c1
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Start property (Publisher)

Returns or sets a **Long** that represents the starting character position of a text range. Read/write.


## Syntax

_expression_.**Start**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

Long


## Remarks

If this property is set to a value larger than that of the **[End](Publisher.TextRange.End.md)** property, the **End** property is set to the same value as that of the **Start** property.


## Example

This example makes the first 15 characters of the selected text range bold. This example assumes that text is selected in the active publication.

```vb
Sub SetSelectionRange() 
 With Selection 
 With .TextRange 
 .Start = 0 
 .End = 15 
 .Font.Bold = msoTrue 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]