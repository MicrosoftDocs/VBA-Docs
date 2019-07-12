---
title: ParagraphFormat object (PowerPoint)
keywords: vbapp10.chm576000
f1_keywords:
- vbapp10.chm576000
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat
ms.assetid: 15d495cf-16e2-5cfb-e99c-a551876e3a8a
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat object (PowerPoint)

Represents the paragraph formatting of a text range.


## Example

Use the [ParagraphFormat](PowerPoint.TextRange.ParagraphFormat.md)property to return the  **ParagraphFormat** object. The following example left aligns the paragraphs in shape two on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _

    .ParagraphFormat.Alignment = ppAlignLeft
```


## Properties



|Name|
|:-----|
|[Alignment](PowerPoint.ParagraphFormat.Alignment.md)|
|[Application](PowerPoint.ParagraphFormat.Application.md)|
|[BaseLineAlignment](PowerPoint.ParagraphFormat.BaseLineAlignment.md)|
|[Bullet](PowerPoint.ParagraphFormat.Bullet.md)|
|[FarEastLineBreakControl](PowerPoint.ParagraphFormat.FarEastLineBreakControl.md)|
|[HangingPunctuation](PowerPoint.ParagraphFormat.HangingPunctuation.md)|
|[LineRuleAfter](PowerPoint.ParagraphFormat.LineRuleAfter.md)|
|[LineRuleBefore](PowerPoint.ParagraphFormat.LineRuleBefore.md)|
|[LineRuleWithin](PowerPoint.ParagraphFormat.LineRuleWithin.md)|
|[Parent](PowerPoint.ParagraphFormat.Parent.md)|
|[SpaceAfter](PowerPoint.ParagraphFormat.SpaceAfter.md)|
|[SpaceBefore](PowerPoint.ParagraphFormat.SpaceBefore.md)|
|[SpaceWithin](PowerPoint.ParagraphFormat.SpaceWithin.md)|
|[TextDirection](PowerPoint.ParagraphFormat.TextDirection.md)|
|[WordWrap](PowerPoint.ParagraphFormat.WordWrap.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]