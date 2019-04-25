---
title: FillFormat.Solid method (Word)
keywords: vbawd10.chm164102159
f1_keywords:
- vbawd10.chm164102159
ms.prod: word
api_name:
- Word.FillFormat.Solid
ms.assetid: 320f5475-7283-c394-0987-3eba3e1d0447
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.Solid method (Word)

Sets the specified fill to a uniform color. 


## Syntax

_expression_.**Solid**

_expression_ Required. A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Remarks

Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.


## Example

This example converts all fills on the active document to uniform red fills.


```vb
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 With shapeLoop.Fill 
 .Solid 
 .ForeColor.RGB = RGB(255, 0, 0) 
 End With 
Next
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]