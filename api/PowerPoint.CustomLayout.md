---
title: CustomLayout object (PowerPoint)
keywords: vbapp10.chm672000
f1_keywords:
- vbapp10.chm672000
ms.prod: powerpoint
api_name:
- PowerPoint.CustomLayout
ms.assetid: 67829704-0314-aed2-5415-6736cefc197e
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomLayout object (PowerPoint)

Represents a custom layout associated with a presentation design. The **CustomLayout** object is a member of the **[CustomLayouts](PowerPoint.CustomLayouts.md)** collection.


## Remarks

Use the  **CustomLayout** property of the **[Slide](PowerPoint.Slide.md)** or **[SlideRange](PowerPoint.SlideRange.md)** objects to access a **CustomLayout** object, for example:


```vb
ActiveWindow.Selection.SlideRange(1).CustomLayout
```


```vb
ActivePresentation.Slides(1).CustomLayout
```

Use the  **[Add](PowerPoint.CustomLayouts.Add.md)** method of the **CustomLayouts** collection to add a new custom layout to the presentation design's custom layouts. Use the **[Item](PowerPoint.CustomLayouts.Add.md)** method to refer to a custom layout. Use the **[Paste](PowerPoint.CustomLayouts.Paste.md)** method to paste the slides on the Clipboard into a custom layout and add the custom layout to the **CustomLayouts** collection.


## Methods



|Name|
|:-----|
|[Copy](PowerPoint.CustomLayout.Copy.md)|
|[Cut](PowerPoint.CustomLayout.Cut.md)|
|[Delete](PowerPoint.CustomLayout.Delete.md)|
|[Duplicate](PowerPoint.CustomLayout.Duplicate.md)|
|[MoveTo](PowerPoint.CustomLayout.MoveTo.md)|
|[Select](PowerPoint.CustomLayout.Select.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.CustomLayout.Application.md)|
|[Background](PowerPoint.CustomLayout.Background.md)|
|[CustomerData](PowerPoint.CustomLayout.CustomerData.md)|
|[Design](PowerPoint.CustomLayout.Design.md)|
|[DisplayMasterShapes](PowerPoint.CustomLayout.DisplayMasterShapes.md)|
|[FollowMasterBackground](PowerPoint.CustomLayout.FollowMasterBackground.md)|
|[Guides](PowerPoint.customlayout.guides.md)|
|[HeadersFooters](PowerPoint.CustomLayout.HeadersFooters.md)|
|[Height](PowerPoint.CustomLayout.Height.md)|
|[Hyperlinks](PowerPoint.CustomLayout.Hyperlinks.md)|
|[Index](PowerPoint.CustomLayout.Index.md)|
|[MatchingName](PowerPoint.CustomLayout.MatchingName.md)|
|[Name](PowerPoint.CustomLayout.Name.md)|
|[Parent](PowerPoint.CustomLayout.Parent.md)|
|[Preserved](PowerPoint.CustomLayout.Preserved.md)|
|[Shapes](PowerPoint.CustomLayout.Shapes.md)|
|[SlideShowTransition](PowerPoint.CustomLayout.SlideShowTransition.md)|
|[ThemeColorScheme](PowerPoint.CustomLayout.ThemeColorScheme.md)|
|[TimeLine](PowerPoint.CustomLayout.TimeLine.md)|
|[Width](PowerPoint.CustomLayout.Width.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]