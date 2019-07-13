---
title: PlaySettings.Application property (PowerPoint)
keywords: vbapp10.chm568001
f1_keywords:
- vbapp10.chm568001
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.Application
ms.assetid: 740e7607-d3a0-b9f6-e13d-9ba3319ecceb
ms.date: 06/08/2017
localization_priority: Normal
---


# PlaySettings.Application property (PowerPoint)

Returns an  **[Application](PowerPoint.Application.md)** object that represents the creator of the specified object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a [PlaySettings](PowerPoint.PlaySettings.md) object.


## Return value

Application


## Example

In this example, a  **[Presentation](PowerPoint.Presentation.md)** object is passed to the procedure. The procedure adds a slide to the presentation and then saves the presentation in the folder where Microsoft PowerPoint is running.


```vb
Sub AddAndSave(pptPres As Presentation)

    pptPres.Slides.Add 1, 1

    pptPres.SaveAs pptPres.Application.Path & "\Added Slide"

End Sub
```

This example displays the name of the application that created each linked OLE object on slide one in the active presentation.




```vb
For Each shpOle In ActivePresentation.Slides(1).Shapes

    If shpOle.Type = msoLinkedOLEObject Then

        MsgBox shpOle.OLEFormat.Application.Name

    End If

Next
```


## See also


[PlaySettings Object](PowerPoint.PlaySettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]