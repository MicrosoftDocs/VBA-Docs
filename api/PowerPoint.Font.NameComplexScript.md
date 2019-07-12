---
title: Font.NameComplexScript property (PowerPoint)
keywords: vbapp10.chm575020
f1_keywords:
- vbapp10.chm575020
ms.prod: powerpoint
api_name:
- PowerPoint.Font.NameComplexScript
ms.assetid: ef1e44d6-ff5d-aaa9-4eaa-643cb2ebc2bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.NameComplexScript property (PowerPoint)

Returns or sets the complex script font name. Used for mixed language text. Read/write.


## Syntax

_expression_. `NameComplexScript`

_expression_ A variable that represents a [Font](PowerPoint.Font.md) object.


## Return value

String


## Remarks

When you have a right-to-left language setting specified, this property is equivalent to the  **Complex scripts font** list in the **Font** dialog box (**Font** tab). When you have an Asian language setting specified, this property is equivalent to the **Asian text font** list in the **Font** dialog box (**Font** tab).


## Example

This example sets the complex script font to Times New Roman.


```vb
ActivePresentation.Slides(1).Shapes.Title.TextFrame _
    .TextRange.Font.NameComplexScript = "Times New Roman"
```


## See also


[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]