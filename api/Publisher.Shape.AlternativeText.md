---
title: Shape.AlternativeText property (Publisher)
keywords: vbapb10.chm2228320
f1_keywords:
- vbapb10.chm2228320
ms.prod: publisher
api_name:
- Publisher.Shape.AlternativeText
ms.assetid: 13bc57af-7067-d60c-5096-a68b1f821d58
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.AlternativeText property (Publisher)

Returns or sets a **String** representing the text displayed by a web browser in place of the **Shape** object while the **Shape** object is being downloaded or when graphics are turned off. Read/write.


## Syntax

_expression_.**AlternativeText**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Remarks

The maximum length of the **AlternativeText** property is 254 characters. Microsoft Publisher returns an error if the text length exceeds this number.


## Example

This example sets the alternative text for the selected shape in the active document. This example assumes that you have a publication in which the selected shape is a picture of a duck.

```vb
Public Sub Alternative_Text() 
 
 ' The picture of a duck must be selected. 
 Publisher.ActiveDocument.Selection.ShapeRange _ 
 .AlternativeText = "This is a mallard duck." 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]