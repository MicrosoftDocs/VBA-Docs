---
title: Presentation.LayoutDirection property (PowerPoint)
keywords: vbapp10.chm583028
f1_keywords:
- vbapp10.chm583028
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.LayoutDirection
ms.assetid: 180e6c85-618f-47e4-b0e7-f9ee3f331c25
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.LayoutDirection property (PowerPoint)

Returns or sets the layout direction for the user interface. Read/write.


## Syntax

_expression_. `LayoutDirection`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

PpDirection


## Remarks

The value of the  **LayoutDirection** property can be one of these **PpDirection** constants. The default value depends on the language support you have selected or installed.


||
|:-----|
|**ppDirectionLeftToRight**|
|**ppDirectionMixed**|
|**ppDirectionRightToLeft**|

## Example

This example sets the layout direction to right-to-left.


```vb
Application.ActivePresentation.LayoutDirection = ppDirectionRightToLeft
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]