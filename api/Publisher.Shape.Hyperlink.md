---
title: Shape.Hyperlink property (Publisher)
keywords: vbapb10.chm2228323
f1_keywords:
- vbapb10.chm2228323
ms.prod: publisher
api_name:
- Publisher.Shape.Hyperlink
ms.assetid: 0990ab32-b4a3-6c89-cb9f-8f8c64ef804f
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Hyperlink property (Publisher)

Returns a **[Hyperlink](Publisher.Hyperlink.md)** object representing the hyperlink associated with the specified shape.


## Syntax

_expression_.**Hyperlink**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example sets shape one on page one in the active publication to jump to the specified website when the shape is chosen.

```vb
Dim hypTemp As Hyperlink 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
hypTemp.Address = "https://www.tailspintoys.com/"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]