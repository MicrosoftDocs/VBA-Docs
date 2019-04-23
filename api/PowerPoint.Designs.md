---
title: Designs object (PowerPoint)
keywords: vbapp10.chm643000
f1_keywords:
- vbapp10.chm643000
ms.prod: powerpoint
api_name:
- PowerPoint.Designs
ms.assetid: 9b02ed6d-9a84-3464-5669-f614e0f33b10
ms.date: 06/08/2017
localization_priority: Normal
---


# Designs object (PowerPoint)

Represents a collection of slide design templates.


## Remarks

Use the [Designs](PowerPoint.Slide.Design.md)property of the  **[Presentation](PowerPoint.Presentation.md)** object to reference a design template.

To add or clone an individual design template, use the  **Designs** collection's[Add](PowerPoint.Designs.Add.md) or [Clone](PowerPoint.Designs.Clone.md)methods, respectively. To refer to an individual design template, use the [Item](PowerPoint.Designs.Item.md)method.

To load a design template, use the [Load](PowerPoint.Designs.Load.md)method.


## Example

The following example adds a new design template to the  **Designs** collection and confirms it was added correctly.


```vb
Sub AddDesignMaster()

    With ActivePresentation.Designs

        .Add designName:="MyDesignName"

        MsgBox .Item("MyDesignName").Name

    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]