---
title: AutoCorrect.DisplayAutoLayoutOptions property (PowerPoint)
keywords: vbapp10.chm666002
f1_keywords:
- vbapp10.chm666002
ms.prod: powerpoint
api_name:
- PowerPoint.AutoCorrect.DisplayAutoLayoutOptions
ms.assetid: 2afaf8e2-a30d-1076-3e78-2ee9a4533482
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.DisplayAutoLayoutOptions property (PowerPoint)

Determines whether Microsoft PowerPoint should display the  **AutoLayout Options** button. Read/write.


## Syntax

_expression_. `DisplayAutoLayoutOptions`

_expression_ A variable that represents an [AutoCorrect](PowerPoint.AutoCorrect.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **DisplayAutoLayoutOptions** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Do not display the  **AutoLayout Options** button.|
|**msoTrue**| Display the **AutoLayout Options** button.|

## Example

This example disables display of the  **AutoCorrect Options** and **AutoLayout Options** buttons.


```vb
Sub HideAutoCorrectOpButton()

    With Application.AutoCorrect

        .DisplayAutoLayoutOptions = msoFalse

        .DisplayAutoLayoutOptions = msoFalse

    End With

End Sub
```


## See also


[AutoCorrect Object](PowerPoint.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]