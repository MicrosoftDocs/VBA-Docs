---
title: TextFrame.VerticalTextAlignment property (Publisher)
keywords: vbapb10.chm3866660
f1_keywords:
- vbapb10.chm3866660
ms.prod: publisher
api_name:
- Publisher.TextFrame.VerticalTextAlignment
ms.assetid: cd809f00-b092-c483-fe99-2aa8043fb684
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.VerticalTextAlignment property (Publisher)

Returns or sets a **[PbVerticalTextAlignmentType](publisher.pbverticaltextalignmenttype.md)** constant that represents the vertical alignment of text in a text box. Read/write.


## Syntax

_expression_.**VerticalTextAlignment**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Remarks

The **VerticalTextAlignment** property value can be one of the **PbVerticalTextAlignmentType** constants.


## Example

This example vertically centers the text in the specified text frame. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub SetVerticalAlignment() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .VerticalTextAlignment = pbVerticalTextAlignmentCenter 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]