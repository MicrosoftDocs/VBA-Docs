---
title: Options object (PowerPoint)
keywords: vbapp10.chm667000
f1_keywords:
- vbapp10.chm667000
ms.prod: powerpoint
api_name:
- PowerPoint.Options
ms.assetid: c129bafc-9927-0171-769e-21649ead7dca
ms.date: 06/08/2017
localization_priority: Normal
---


# Options object (PowerPoint)

Represents application options in Microsoft PowerPoint.


## Example

Use the  **[Options](PowerPoint.Application.Options.md)** property to return an **Options** object. The following example sets three application options for PowerPoint.


```vb
Sub TogglePasteOptionsButton()

    With Application.Options

        If .DisplayPasteOptions = False Then

            .DisplayPasteOptions = True

        End If

    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]