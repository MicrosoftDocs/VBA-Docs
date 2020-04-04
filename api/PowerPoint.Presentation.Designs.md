---
title: Presentation.Designs property (PowerPoint)
keywords: vbapp10.chm583063
f1_keywords:
- vbapp10.chm583063
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Designs
ms.assetid: 5ad47ac9-aaab-3971-1102-fa48e8bcef8b
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Designs property (PowerPoint)

Returns a **[Designs](PowerPoint.Designs.md)** object that represents a collection of designs.


## Syntax

_expression_. `Designs`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Designs


## Example

The following example displays a message for each design in the active presentation.


```vb
Sub AddDesignMaster()

    Dim desName As Design



    With ActivePresentation



        For Each desName In .Designs

            MsgBox "The design name is " & .Designs.Item(desName.Index).Name

        Next



    End With



End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]