---
title: Cell.VerticalTextAlignment property (Publisher)
keywords: vbapb10.chm5111840
f1_keywords:
- vbapb10.chm5111840
ms.prod: publisher
api_name:
- Publisher.Cell.VerticalTextAlignment
ms.assetid: 793bf932-15d0-cce9-1d5d-aee5d260e1a0
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.VerticalTextAlignment property (Publisher)

Returns or sets a **[PbVerticalTextAlignmentType](publisher.pbverticaltextalignmenttype.md)** constant that represents the vertical alignment of text in a text box. Read/write.


## Syntax

_expression_.**VerticalTextAlignment**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


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