---
title: TextRange2.Font property (PowerPoint)
ms.assetid: 3d47ff57-6622-4eaa-b8ff-b395e9757096
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.Font property (PowerPoint)

Returns a  **Font** object that represents character formatting for the **TextRange2** object. Read-only.


## Syntax

_expression_.**Font**

 _expression_ An expression that returns a 'TextRange2' object.


## Return value

Font


## Example

This example sets the formatting for the text in shape one on slide one in the active PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 With .TextFrame.TextRange2.Font 
 .Size = 48 
 .Name = "Palatino" 
 .Bold = True 
 .Color.RGB = RGB(255, 127, 255) 
 End With 
End With
```


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]