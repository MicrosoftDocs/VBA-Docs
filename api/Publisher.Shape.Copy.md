---
title: Shape.Copy method (Publisher)
keywords: vbapb10.chm2228242
f1_keywords:
- vbapb10.chm2228242
ms.prod: publisher
api_name:
- Publisher.Shape.Copy
ms.assetid: cfec06d8-9f9b-4d88-eb28-e9e29fb1aed1
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Copy method (Publisher)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


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