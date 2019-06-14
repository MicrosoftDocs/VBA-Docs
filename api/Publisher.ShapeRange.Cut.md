---
title: ShapeRange.Cut method (Publisher)
keywords: vbapb10.chm2293777
f1_keywords:
- vbapb10.chm2293777
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Cut
ms.assetid: 961d4646-8318-d2ff-ed98-649583d36115
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Cut method (Publisher)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

Use the **[Shapes.Paste](Publisher.Shapes.Paste.md)** method to paste the contents of the Clipboard.

The **Copy** method can be used on **Shape** objects, but the **Paste** method cannot.


## Example

This example deletes shape one and shape two from page one of the active publication, places copies of them on the Clipboard, and then pastes the copies onto page two.

```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Cut 
 .Pages(2).Shapes.Paste 
End With
```

<br/>

This example deletes shape one on page one of the active publication and places a copy of it on the Clipboard.

```vb
ActiveDocument.Pages(1).Shapes(1).Cut

```

<br/>

This example deletes the text in shape one on page one of the active publication and places a copy of it on the Clipboard.

```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Cut
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]