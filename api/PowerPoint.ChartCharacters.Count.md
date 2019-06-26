---
title: ChartCharacters.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartCharacters.Count
ms.assetid: 99e1634b-49de-220e-e0e1-cfb31a1ba73a
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartCharacters.Count property (PowerPoint)

Returns the number of objects in the collection. Read-only  **Long**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a '[ChartCharacters](PowerPoint.ChartCharacters.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example makes the last character a superscript character in the title of the first chart in the active document.




```vb
Sub MakeSuperscript()

    Dim n As Integer



    With ActiveDocument.InlineShapes(1)

        If .HasChart Then

            n = .Chart.Title.Characters.Count

            .Chart.Title.Characters(n, 1).Font.Superscript = True

        End If

    End With

End Sub
```


## See also


[ChartCharacters Object](PowerPoint.ChartCharacters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]