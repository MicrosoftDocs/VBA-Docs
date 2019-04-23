---
title: Shape.Cut method (Publisher)
keywords: vbapb10.chm2228241
f1_keywords:
- vbapb10.chm2228241
ms.prod: publisher
api_name:
- Publisher.Shape.Cut
ms.assetid: d800c1e5-7655-9071-a373-7772fa1ca15f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Cut method (Publisher)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

 _expression_ A variable that represents a  **Shape** object.


## Remarks

Use the  **Paste**method to paste the contents of the Clipboard.

The  **Copy** method can be used on **Shape** objects, but the **Paste** method cannot.


## Example

This example deletes shape one and shape two from page one of the active publication, places copies of them on the Clipboard, and then pastes the copies onto page two.


```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Cut 
 .Pages(2).Shapes.Paste 
End With
```

This example deletes shape one on page one of the active publication and places a copy of it on the Clipboard.




```vb
ActiveDocument
```




```vb
.Pages(1).Shapes(1).Cut
```

This example deletes the text in shape one on page one of the active publication and places a copy of it on the Clipboard.




```vb
ActiveDocument
```




```vb
.Pages(1).Shapes(1).TextFrame.TextRange.Cut
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]