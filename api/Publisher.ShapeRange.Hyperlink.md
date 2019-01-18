---
title: ShapeRange.Hyperlink Property (Publisher)
keywords: vbapb10.chm2293859
f1_keywords:
- vbapb10.chm2293859
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Hyperlink
ms.assetid: 34ec968c-af66-7629-066f-80c8e1b40e84
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Hyperlink Property (Publisher)

Returns a  **[Hyperlink](Publisher.Hyperlink.md)** object representing the hyperlink associated with the specified shape.


## Syntax

 _expression_. **Hyperlink**

 _expression_ A variable that represents a  **ShapeRange** object.


## Example

This example sets shape one on page one in the active publication to jump to the specified Web site when the shape is clicked.


```vb
Dim hypTemp As Hyperlink 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
hypTemp.Address = "https://www.tailspintoys.com/"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]