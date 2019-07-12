---
title: Master object (PowerPoint)
keywords: vbapp10.chm638000
f1_keywords:
- vbapp10.chm638000
ms.prod: powerpoint
api_name:
- PowerPoint.Master
ms.assetid: 22e8805e-6469-1a34-7f7b-f1ea5c6c49ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Master object (PowerPoint)

Represents a slide master, title master, handout master, notes master, or design master.


## Example

To return a  **Master** object, use the [Master](PowerPoint.Slide.Master.md)property of the  **[Slide](PowerPoint.Slide.md)** object or **[SlideRange](PowerPoint.SlideRange.md)** collection, or use the [HandoutMaster](PowerPoint.Presentation.HandoutMaster.md), [NotesMaster](PowerPoint.Presentation.NotesMaster.md), [SlideMaster](PowerPoint.Design.SlideMaster.md), or [TitleMaster](PowerPoint.Presentation.TitleMaster.md)property of the  **[Presentation](PowerPoint.Presentation.md)** object. Note that some of these properties are also available from the **[Design](PowerPoint.Design.md)** object as well. The following example sets the background fill for the slide master for the active presentation.


```vb
ActivePresentation.SlideMaster.Background.Fill _

    .PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```

To add a title master or design to a presentation and return a  **Master** object that represents the new title master or design, use the [AddTitleMaster](PowerPoint.Presentation.AddTitleMaster.md)method. The following example adds a title master to the active presentation and places the title placeholder 10 points from the top of the master.




```vb
ActivePresentation.AddTitleMaster.Shapes.Title.Top = 10
```


## Methods



|Name|
|:-----|
|[ApplyTheme](PowerPoint.Master.ApplyTheme.md)|
|[Delete](PowerPoint.Master.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Master.Application.md)|
|[Background](PowerPoint.Master.Background.md)|
|[BackgroundStyle](PowerPoint.Master.BackgroundStyle.md)|
|[ColorScheme](PowerPoint.Master.ColorScheme.md)|
|[CustomerData](PowerPoint.Master.CustomerData.md)|
|[CustomLayouts](PowerPoint.Master.CustomLayouts.md)|
|[Design](PowerPoint.Master.Design.md)|
|[Guides](PowerPoint.master.guides.md)|
|[HeadersFooters](PowerPoint.Master.HeadersFooters.md)|
|[Height](PowerPoint.Master.Height.md)|
|[Hyperlinks](PowerPoint.Master.Hyperlinks.md)|
|[Name](PowerPoint.Master.Name.md)|
|[Parent](PowerPoint.Master.Parent.md)|
|[Shapes](PowerPoint.Master.Shapes.md)|
|[SlideShowTransition](PowerPoint.Master.SlideShowTransition.md)|
|[TextStyles](PowerPoint.Master.TextStyles.md)|
|[Theme](PowerPoint.Master.Theme.md)|
|[TimeLine](PowerPoint.Master.TimeLine.md)|
|[Width](PowerPoint.Master.Width.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]