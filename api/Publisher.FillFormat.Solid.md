---
title: FillFormat.Solid method (Publisher)
keywords: vbapb10.chm2359317
f1_keywords:
- vbapb10.chm2359317
ms.prod: publisher
api_name:
- Publisher.FillFormat.Solid
ms.assetid: e34f6bc0-308b-4f86-5ce9-87e05c4a2089
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.Solid method (Publisher)

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.


## Syntax

_expression_.**Solid**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Example

This example converts all fills on the first page of the active publication to uniform red fills.

```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 .Solid 
 .ForeColor.RGB = RGB(255, 0, 0) 
 End With 
Next shpLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]