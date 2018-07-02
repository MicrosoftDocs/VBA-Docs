---
title: TextFrame2.DeleteText Method (Office)
ms.prod: office
api_name:
- Office.TextFrame2.DeleteText
ms.assetid: 4bfd3a9b-e902-0f83-f1fe-19dd95115278
ms.date: 06/08/2017
---


# TextFrame2.DeleteText Method (Office)

Deletes the text from a text frame and all the associated properties of the text, including font attributes.


## Syntax

 _expression_. `DeleteText`

 _expression_ An expression that returns a [TextFrame2](./Office.TextFrame2.md) object.


### Return Value

Nothing


## Example

The following code shows how to delete the text from shape one on slide one of the active presentation, if that shape contains text.


```vb
Dim pptSlide As Slide 
Set pptSlide = ActivePresentation.Slides(1) 
pptSlide.Shapes(1).TextFrame2.DeleteText
```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/textframe2-members-office.md)

