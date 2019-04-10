---
title: Model3DFormat.Application property (Word)
keywords: vbawd10.chm151585768
f1_keywords:
- vbawd10.chm151585768
ms.prod: word
api_name:
- Word.Model3DFormat.Application
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.Application property (Word)

Returns an **[Application](Word.Application.md)** object that represents the creator of the specified object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Model3DFormat](Word.Model3DFormat.md)** object.


## Return value

Object


## Example

In this example, a **[Presentation](powerpoint.presentation.md)** object is passed to the procedure. The procedure adds a slide to the presentation and then saves the presentation in the folder where Microsoft Word is running.

```vb
Sub AddAndSave(pptPres As Presentation)

    pptPres.Slides.Add 1, 1

    pptPres.SaveAs pptPres.Application.Path & "\Added Slide"

End Sub
```

<br/>

This example displays the name of the application that created each linked OLE object on slide one in the active presentation.

```vb
For Each shpOle In ActivePresentation.Slides(1).Shapes

    If shpOle.Type = msoLinkedOLEObject Then

        MsgBox shpOle.OLEFormat.Application.Name

    End If

Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]