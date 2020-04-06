---
title: AutoCorrect.DisplayAutoCorrectOptions property (PowerPoint)
keywords: vbapp10.chm666001
f1_keywords:
- vbapp10.chm666001
ms.prod: powerpoint
api_name:
- PowerPoint.AutoCorrect.DisplayAutoCorrectOptions
ms.assetid: d3d769aa-af42-27c2-1c8e-39684d4f70a7
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.DisplayAutoCorrectOptions property (PowerPoint)

Determines whether Microsoft PowerPoint should display the  **AutoCorrect Options** button. Read/write.


## Syntax

_expression_. `DisplayAutoCorrectOptions`

_expression_ A variable that represents an [AutoCorrect](PowerPoint.AutoCorrect.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **DisplayAutoCorrectOptions** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Do not display the  **AutoCorrect Options** button.|
|**msoTrue**| Display the **AutoCorrect Options** button.|

## Example

This example disables display of the  **AutoCorrect Options** and **AutoLayout Options** buttons.


```vb
Sub HideAutoCorrectOpButton()

    With Application.AutoCorrect

        .DisplayAutoCorrectOptions = msoFalse

        .DisplayAutoLayoutOptions = msoFalse

    End With

End Sub
```


## See also


[AutoCorrect Object](PowerPoint.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]