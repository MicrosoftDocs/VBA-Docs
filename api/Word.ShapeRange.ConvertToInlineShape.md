---
title: ShapeRange.ConvertToInlineShape method (Word)
keywords: vbawd10.chm162856990
f1_keywords:
- vbawd10.chm162856990
ms.prod: word
api_name:
- Word.ShapeRange.ConvertToInlineShape
ms.assetid: 01ce99b9-408b-2bd4-fd05-21d17e2ada91
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.ConvertToInlineShape method (Word)

Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only shapes that represent pictures, OLE objects, or ActiveX controls. .


## Syntax

_expression_. `ConvertToInlineShape`

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Return value

 **[InlineShape](Word.InlineShape.md)**


## Remarks

Shapes that support attached text cannot be converted to inline shapes. For these shapes, use the **ConvertToFrame** method.

If you use this method on a  **ShapeRange** object that contains more than one shape, an error occurs.


## Example

This example converts each picture in MyDoc.doc to an inline shape.


```vb
For Each s In Documents("MyDoc.doc").Shapes 
 If s.Type = msoPicture Then 
 s.ConvertToInlineShape 
 End If 
Next s
```


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]