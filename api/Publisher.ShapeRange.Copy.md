---
title: ShapeRange.Copy method (Publisher)
keywords: vbapb10.chm2293778
f1_keywords:
- vbapb10.chm2293778
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Copy
ms.assetid: 11b9da00-85e4-fc7a-fa93-4a451b7bd15a
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Copy method (Publisher)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Return value

Nothing


## Remarks

Use the **[Shapes.Paste](Publisher.Shapes.Paste.md)** method to paste the contents of the Clipboard.

The **Copy** method can be used on **Shape** objects, but the **Paste** method cannot.


## Example

This example copies shapes one and two on page one of the active publication to the Clipboard, and then pastes the copies onto page two.

```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Copy 
 .Pages(2).Shapes.Paste 
End With
```

<br/>

This example copies shape one on page one of the active publication to the Clipboard.

```vb
ActiveDocument.Pages(1).Shapes(1).Copy
```

<br/>

This example copies the text in shape one on page one of the active publication to the Clipboard.

```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Copy
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]