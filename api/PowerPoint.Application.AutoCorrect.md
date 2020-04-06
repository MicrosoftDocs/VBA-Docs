---
title: Application.AutoCorrect property (PowerPoint)
keywords: vbapp10.chm502053
f1_keywords:
- vbapp10.chm502053
ms.prod: powerpoint
api_name:
- PowerPoint.Application.AutoCorrect
ms.assetid: 490fc728-c639-2a32-22b8-1757c14e9bd7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AutoCorrect property (PowerPoint)

Returns an  **[AutoCorrect](PowerPoint.AutoCorrect.md)** object that represents the AutoCorrect functionality in Microsoft PowerPoint.


## Syntax

_expression_. `AutoCorrect`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

AutoCorrect


## Example

The following example disables display of the  **AutoCorrect Options** and **AutoLayout Options** buttons.


```vb
Sub HideAutoCorrectOpButton()

    With Application.AutoCorrect

        .DisplayAutoCorrectOptions = msoFalse

        .DisplayAutoLayoutOptions = msoFalse

    End With

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]