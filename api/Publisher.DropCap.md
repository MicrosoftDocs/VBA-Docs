---
title: DropCap object (Publisher)
keywords: vbapb10.chm5570559
f1_keywords:
- vbapb10.chm5570559
ms.prod: publisher
api_name:
- Publisher.DropCap
ms.assetid: 7c6aeffe-cf25-a834-52de-5966df5e21d2
ms.date: 05/31/2019
localization_priority: Normal
---


# DropCap object (Publisher)

Represents a dropped capital letter at the beginning of a paragraph.
 
## Remarks

Use the **[DropCap](Publisher.TextRange.DropCap.md)** property of the **TextRange** object to return a **DropCap** object. 

## Example

The following example sets a dropped capital letter for the first letter of each paragraph in the first shape on the first page of the active publication. This example assumes that the specified shape is a text box and not another type of shape.

```vb
Sub ApplyDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .DropCap.ApplyCustomDropCap Size:=3, Span:=3, Bold:=True 
End Sub
```


## Methods

- [ApplyCustomDropCap](Publisher.DropCap.ApplyCustomDropCap.md)
- [Clear](Publisher.DropCap.Clear.md)

## Properties

- [Application](Publisher.DropCap.Application.md)
- [FontBold](Publisher.DropCap.FontBold.md)
- [FontColor](Publisher.DropCap.FontColor.md)
- [FontItalic](Publisher.DropCap.FontItalic.md)
- [FontName](Publisher.DropCap.FontName.md)
- [LinesUp](Publisher.DropCap.LinesUp.md)
- [Parent](Publisher.DropCap.Parent.md)
- [Size](Publisher.DropCap.Size.md)
- [Span](Publisher.DropCap.Span.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]