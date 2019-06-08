---
title: Hyperlink.Shape property (Publisher)
keywords: vbapb10.chm4587527
f1_keywords:
- vbapb10.chm4587527
ms.prod: publisher
api_name:
- Publisher.Hyperlink.Shape
ms.assetid: afd1dab7-472a-2aa5-f5da-1e2f783b5270
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.Shape property (Publisher)

Returns a **[Shape](Publisher.Shape.md)** object that represents the shape associated with a hyperlink.


## Syntax

_expression_.**Shape**

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


## Return value

Shape


## Example

This example adds a hyperlink to the first shape on the first page of the active publication and then vertically flips the shape. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub FormatHyperlinkShape() 
 With ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 .Address = "https://www.tailspintoys.com/" 
 .Shape.Flip FlipCmd:=msoFlipVertical 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]