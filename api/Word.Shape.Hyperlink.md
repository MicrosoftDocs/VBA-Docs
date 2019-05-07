---
title: Shape.Hyperlink property (Word)
keywords: vbawd10.chm161481705
f1_keywords:
- vbawd10.chm161481705
ms.prod: word
api_name:
- Word.Shape.Hyperlink
ms.assetid: bd69df23-a1b0-cdce-64a4-b0b6046e96d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Hyperlink property (Word)

Returns a  **Hyperlink** object that represents the hyperlink associated with a **Shape** object. Read-only.


## Syntax

_expression_.**Hyperlink**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Remarks

If there is no hyperlink associated with the specified shape, an error occurs. In this case, use the  **[Add](Word.Hyperlinks.Add.md)** method for the **[Hyperlinks](Word.hyperlinks.md)** collection to add a hyperlink to the specified shape. The following example shows how to do this.


```vb
ActiveDocument.Hyperlinks.Add Selection.Shapes(1), "https://www.microsoft.com"
```


## Example

This example displays the address for the hyperlink for the first shape in the active document.


```vb
MsgBox ActiveDocument.Shapes(1).Hyperlink.Address
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]