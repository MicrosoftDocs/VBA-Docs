---
title: PrintOptions.SlideShowName property (PowerPoint)
keywords: vbapp10.chm517014
f1_keywords:
- vbapp10.chm517014
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.SlideShowName
ms.assetid: 0eca6fce-23ba-0e13-a6a7-ee968f3aa973
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintOptions.SlideShowName property (PowerPoint)

Returns or sets the name of the custom slide show to print. Read/write .


## Syntax

_expression_. `SlideShowName`

_expression_ A variable that represents a [PrintOptions](PowerPoint.PrintOptions.md) object.


## Return value

String


## Remarks

To print a custom slide show, you must first set the  **[RangeType](PowerPoint.PrintOptions.RangeType.md)** property to **ppPrintNamedSlideShow**.


## Example

This example prints an existing custom slide show named "tech talk."


```vb
With ActivePresentation.PrintOptions

    .RangeType = ppPrintNamedSlideShow

    .SlideShowName = "tech talk"

End With

ActivePresentation.PrintOut
```


## See also


[PrintOptions Object](PowerPoint.PrintOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]