---
title: Hyperlink.TextToDisplay property (Publisher)
keywords: vbapb10.chm4587536
f1_keywords:
- vbapb10.chm4587536
ms.prod: publisher
api_name:
- Publisher.Hyperlink.TextToDisplay
ms.assetid: 26b5857c-3f94-0d33-f65e-9c34f2a4cc2b
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.TextToDisplay property (Publisher)

Returns or sets a **String** that represents the text displayed for a hyperlink. Read/write.


## Syntax

_expression_.**TextToDisplay**

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


## Return value

String


## Example

This example sets the hyperlink display text and address of the first hyperlink on the first page. This example assumes that the first page of the active publication contains at least one shape with at least one text hyperlink.

```vb
Sub SetHyperlinkTextToDisplay() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1) 
 .TextToDisplay = "Tailspin Toys website" 
 .Address = "https://www.tailspintoys.com/" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]