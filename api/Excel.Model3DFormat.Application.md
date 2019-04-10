---
title: Model3DFormat.Application property (Excel)
ms.prod: excel
api_name:
- Excel.Model3DFormat.Application
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.Application property (Excel)

Returns an **[Application](excel.application(object).md)** object that represents the creator of the specified object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Model3DFormat](Excel.Model3DFormat.md)** object.


## Return value

Object


## Example

In this example, a **[Presentation](powerpoint.presentation.md)** object is passed to the procedure. The procedure adds a slide to the presentation and then saves the presentation in the folder where Microsoft Excel is running.

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