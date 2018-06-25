---
title: TextFrame2.WordArtformat Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.WordArtformat
ms.assetid: b9d6c36d-e353-940f-4984-1f5ed3cf165c
ms.date: 06/08/2017
---


# TextFrame2.WordArtformat Property (Office)

Returns or sets the WordArt type for the specified text frame. Read/write


## Syntax

 _expression_. `WordArtformat`

 _expression_ An expression that returns a [TextFrame2](./Office.TextFrame2.md) object.


## Remarks

The value of the WordArtFormat property can be one of these MsoPresetTextEffect constants.


## Example

The following code shows how to set the word art format for shape one on slide one in the active presentation.


```vb
Dim pptSlide As Slide 
Set pptSlide = ActivePresentation.Slides(1) 
pptSlide.Shapes(1).TextFrame2.WordArtFormat = msoTextEffect20 

```


## See also


[TextFrame2 Object](Office.TextFrame2.md)
#### Other resources


[TextFrame2 Object Members](./overview/textframe2-members-office.md)

