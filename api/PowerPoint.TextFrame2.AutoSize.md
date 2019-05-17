---
title: TextFrame2.AutoSize property (PowerPoint)
keywords: vbapp10.chm678013
f1_keywords:
- vbapp10.chm678013
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.AutoSize
ms.assetid: 48f05f1b-8269-742f-20ca-6ebdde5fa682
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.AutoSize property (PowerPoint)

 Returns or sets a value that indicates whether the size of the specified shape is changed automatically to fit text within its boundaries. Read/write.


## Syntax

_expression_.**AutoSize**

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

MsoAutoSize


## Remarks

The value of the  **AutoSize** property can be one of the following **MsoAutoSize** constants.


||
|:-----|
|**msoAutoSizeMixed**|
|**msoAutoSizeNone**|
|**msoAutoSizeShapeToFitText**|
|**msoAutoSizeTextToFitShape**|

## Example

The following example shows how to adjust the size of the title text on slide one to fit the text frame that contains it.


```vb
Public Sub AutoSize_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes(1)

        If .TextFrame2.TextRange.Characters.Count < 50 Then

            .TextFrame2.AutoSize = msoAutoSizeTextToFitShape

        End If

    End With

    

End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]