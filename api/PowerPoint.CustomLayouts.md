---
title: CustomLayouts object (PowerPoint)
keywords: vbapp10.chm671000
f1_keywords:
- vbapp10.chm671000
ms.prod: powerpoint
api_name:
- PowerPoint.CustomLayouts
ms.assetid: 9ce682fb-545c-55cb-e9ac-3475f7556af1
ms.date: 12/26/2018
localization_priority: Normal
---


# CustomLayouts object (PowerPoint)

Represents a set of custom layouts associated with a presentation design.

## Remarks

Use the **[CustomLayouts](PowerPoint.Master.CustomLayouts.md)** property of the slide **[Master](PowerPoint.Master.md)** object to return a **CustomLayouts** collection. Use **CustomLayouts** (_index_), where index is the custom layout index number, to return a single **[CustomLayout](PowerPoint.CustomLayout.md)** object.

Use the **[Add](PowerPoint.CustomLayouts.Add.md)** method to create a new custom layout and add it to the **CustomLayouts** collection. Use the **[Paste](PowerPoint.CustomLayouts.Paste.md)** method to past slides from the Clipboard as a **CustomLayout** object into the **CustomLayouts** collection.

Use the **CustomLayout** property of a **[Slide](PowerPoint.Slide.md)** or **[SlideRange](PowerPoint.SlideRange.md)** object to return a custom layout for a slide or set of slides.


## Example

The following example adds a custom layout to the slide master of the active presentation.


```vb
Sub AddCustomLayout()

    With ActivePresentation.SlideMaster

        .CustomLayouts.Add (1)

        .CustomLayouts(1).Name = "MyLayout"

    End With

End Sub
```

The following example displays the name of the custom layout for the first slide of the active presentation.

```vb
MsgBox ActivePresentation.Slides(1).CustomLayout.Name
```

## Methods

|Name|
|:-----|
|[Add](PowerPoint.CustomLayouts.Add.md)|
|[Item](PowerPoint.CustomLayouts.Item.md)|
|[Paste](PowerPoint.CustomLayouts.Paste.md)|

## Properties

|Name|
|:-----|
|[Application](PowerPoint.CustomLayouts.Application.md)|
|[Count](PowerPoint.CustomLayouts.Count.md)|
|[Parent](PowerPoint.CustomLayouts.Parent.md)|

## See also

- [PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]