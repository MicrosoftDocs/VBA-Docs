---
title: InlineShape.Hyperlink property (Word)
keywords: vbawd10.chm162004999
f1_keywords:
- vbawd10.chm162004999
ms.prod: word
api_name:
- Word.InlineShape.Hyperlink
ms.assetid: 46297480-026a-1679-20dc-f1e6b284559e
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.Hyperlink property (Word)

Returns a  **Hyperlink** object that represents the hyperlink associated with the specified inline shape. Read-only.


## Syntax

_expression_.**Hyperlink**

_expression_ A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Remarks

If there is no hyperlink associated with the specified shape, an error occurs. In this case, use the  **[Add](Word.Hyperlinks.Add.md)** method for the **[Hyperlinks](Word.hyperlinks.md)** collection to add a hyperlink to the specified shape. The following example shows how to do this.


```vb
ActiveDocument.Hyperlinks.Add Selection.InlineShapes(1), "https://www.microsoft.com"
```


## Example

This example displays the address for the hyperlink for the first shape in the active document.


```vb
MsgBox ActiveDocument.Shapes(1).Hyperlink.Address
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]