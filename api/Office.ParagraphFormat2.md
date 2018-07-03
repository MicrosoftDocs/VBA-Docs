---
title: ParagraphFormat2 Object (Office)
ms.prod: office
api_name:
- Office.ParagraphFormat2
ms.assetid: 05ff2b24-9603-f923-d053-e736fb2ba389
ms.date: 06/08/2017
---


# ParagraphFormat2 Object (Office)

Represents the paragraph formatting of a text range.


## Example

The following example left aligns the paragraphs in shape two on slide one in the active PowerPoint presentation.


```vb
ActivePresentation.Slides(1).Shapes(2).TextFrame2.TextRange2 _ 
 .ParagraphFormat2.Alignment = ppAlignLeft 

```


## Properties



|**Name**|
|:-----|
|[Alignment](Office.ParagraphFormat2.Alignment.md)|
|[Application](Office.ParagraphFormat2.Application.md)|
|[BaselineAlignment](Office.ParagraphFormat2.BaselineAlignment.md)|
|[Bullet](Office.ParagraphFormat2.Bullet.md)|
|[Creator](Office.ParagraphFormat2.Creator.md)|
|[FarEastLineBreakLevel](Office.ParagraphFormat2.FarEastLineBreakLevel.md)|
|[FirstLineIndent](Office.ParagraphFormat2.FirstLineIndent.md)|
|[HangingPunctuation](Office.ParagraphFormat2.HangingPunctuation.md)|
|[IndentLevel](Office.ParagraphFormat2.IndentLevel.md)|
|[LeftIndent](Office.ParagraphFormat2.LeftIndent.md)|
|[LineRuleAfter](Office.ParagraphFormat2.LineRuleAfter.md)|
|[LineRuleBefore](Office.ParagraphFormat2.LineRuleBefore.md)|
|[LineRuleWithin](Office.ParagraphFormat2.LineRuleWithin.md)|
|[Parent](Office.ParagraphFormat2.Parent.md)|
|[RightIndent](Office.ParagraphFormat2.RightIndent.md)|
|[SpaceAfter](Office.ParagraphFormat2.SpaceAfter.md)|
|[SpaceBefore](Office.ParagraphFormat2.SpaceBefore.md)|
|[SpaceWithin](Office.ParagraphFormat2.SpaceWithin.md)|
|[TabStops](Office.ParagraphFormat2.TabStops.md)|
|[TextDirection](Office.ParagraphFormat2.TextDirection.md)|
|[WordWrap](Office.ParagraphFormat2.WordWrap.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
