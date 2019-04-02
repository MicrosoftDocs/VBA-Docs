---
title: Ruler object (PowerPoint)
keywords: vbapp10.chm570000
f1_keywords:
- vbapp10.chm570000
ms.prod: powerpoint
api_name:
- PowerPoint.Ruler
ms.assetid: dc6b78ae-4745-0bc8-1d28-831b1f30f86c
ms.date: 06/08/2017
localization_priority: Normal
---


# Ruler object (PowerPoint)

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.


## Example

Use the [Ruler](PowerPoint.TextFrame.Ruler.md)property of the  **TextFrame** object to return the **Ruler** object that represents the ruler for the text in the specified shape. Use the [TabStops](PowerPoint.Ruler.TabStops.md)property to return the  **TabStops** object that contains the tab stops on the ruler. Use the [Levels](PowerPoint.Ruler.Levels.md)property to return the  **RulerLevels** object that contains the indentation settings for text outline levels. The following example sets a left-aligned tab stop at 2 inches (144 Points) and sets a hanging indent for the text in object two on slide one in the active presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2).TextFrame.Ruler

    .TabStops.Add ppTabStopLeft, 144

    .Levels(1).FirstMargin = 0

    .Levels(1).LeftMargin = 36

End With
```

Use the [Ruler](PowerPoint.TextStyle.Ruler.md)property of the  **TextStyle** object to return the **Ruler** object that represents the ruler for one of the four defined text styles (title text, body text, notes text, or default text). The following example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.




```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Ruler.Levels(1)
    .FirstMargin = 9
    .LeftMargin = 54
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]