---
title: AutoCorrect object (PowerPoint)
keywords: vbapp10.chm666000
f1_keywords:
- vbapp10.chm666000
ms.prod: powerpoint
api_name:
- PowerPoint.AutoCorrect
ms.assetid: c7d0c7a5-220e-6290-b326-cb5cf17c458b
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect object (PowerPoint)

Represents the AutoCorrect functionality in Microsoft PowerPoint.


## Example

Use the [AutoCorrect](PowerPoint.Application.AutoCorrect.md)property to return an **AutoCorrect** object. The following example disables displaying the AutoCorrect options buttons.


```vb
Sub HideAutoCorrectOpButton()

    With Application.AutoCorrect

        .DisplayAutoCorrectOptions = msoFalse

        .DisplayAutoLayoutOptions = msoFalse

    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]