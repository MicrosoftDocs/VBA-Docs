---
title: TextFrame2.HasText property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.HasText
ms.assetid: 4783db2d-8dd5-f9d5-5cfd-8e119868c57e
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2.HasText property (Office)

Indicates whether the shape that contains the specified text frame has text associated with it. Read-only.


## Syntax

_expression_.**HasText**

_expression_ An expression that returns a **[TextFrame2](Office.TextFrame2.md)** object.


## Remarks

The value of the **HasText** property can be one of the following **[MsoTriState](office.msotristate.md)** constants.

|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified text frame does not have text.|
|**msoTrue**| The specified text frame has text.|

## Example

The following code tests whether shape two on slide one contains text, and if it does, resizes the shape to fit the text.


```vb
Dim pptSlide As Slide 
Set pptSlide = ActivePresentation.Slides(1) 
 With pptSlide.Shapes(2).TextFrame 
 If .HasText Then .AutoSize = ppAutoSizeShapeToFitText 
 End With
```


## See also

- [TextFrame2 object members](overview/Library-Reference/textframe2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]