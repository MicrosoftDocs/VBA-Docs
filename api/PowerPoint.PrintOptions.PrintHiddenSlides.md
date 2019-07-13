---
title: PrintOptions.PrintHiddenSlides property (PowerPoint)
keywords: vbapp10.chm517009
f1_keywords:
- vbapp10.chm517009
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.PrintHiddenSlides
ms.assetid: 39b5845e-7fd0-6759-bf1c-e2497acc1c61
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintOptions.PrintHiddenSlides property (PowerPoint)

Determines whether hidden slides in the specified presentation will be printed. Read/write.


## Syntax

_expression_. `PrintHiddenSlides`

_expression_ A variable that represents a [PrintOptions](PowerPoint.PrintOptions.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **PrintHiddenSlides** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The hidden slides in the specified presentation will not be printed.|
|**msoTrue**| The hidden slides in the specified presentation will be printed.|

## Example

This example prints all slides, whether visible or hidden, in the active presentation.


```vb
With ActivePresentation

    .PrintOptions.PrintHiddenSlides = msoTrue

    .PrintOut

End With
```


## See also


[PrintOptions Object](PowerPoint.PrintOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]