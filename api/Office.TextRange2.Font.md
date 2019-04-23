---
title: TextRange2.Font property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Font
ms.assetid: 005fa6bf-2dd5-32ec-18e8-30ff6260e55d
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.Font property (Office)

Returns a **Font** object that represents character formatting for the **TextRange2** object. Read-only.


## Syntax

_expression_.**Font**

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


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

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]