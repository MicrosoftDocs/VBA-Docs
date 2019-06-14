---
title: TextRange.Cut method (Publisher)
keywords: vbapb10.chm5308473
f1_keywords:
- vbapb10.chm5308473
ms.prod: publisher
api_name:
- Publisher.TextRange.Cut
ms.assetid: c9b8b896-26e7-ac58-0e1a-a66ef789f397
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Cut method (Publisher)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

Nothing


## Remarks

Use the **[Paste](Publisher.TextRange.Paste.md)** method to paste the contents of the Clipboard.

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