---
title: TextFrame2.WarpFormat property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.WarpFormat
ms.assetid: 83993a3d-a594-e3bc-47ca-47f50be143b7
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2.WarpFormat property (Office)

Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write.


## Syntax

_expression_.**WarpFormat**

_expression_ An expression that returns a **[TextFrame2](Office.TextFrame2.md)** object.


## Remarks

The value of the **WarpFormat** property can be one of the **[MsoWarpFormat](office.msowarpformat.md)** constants.


## Example

The following code shows how to set the warp format for shape one on slide one of the active presentation.


```vb
Public Sub WarpFormat_Example() 
 
 Dim pptSlide As Slide 
 Set pptSlide = ActivePresentation.Slides(1) 
 pptSlide.Shapes(1).TextFrame2.WarpFormat = msoWarpFormat15 
 
End Sub 

```


## See also

- [TextFrame2 object members](overview/Library-Reference/textframe2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]