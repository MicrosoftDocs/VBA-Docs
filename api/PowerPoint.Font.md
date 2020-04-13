---
title: Font object (PowerPoint)
keywords: vbapp10.chm575000
f1_keywords:
- vbapp10.chm575000
ms.prod: powerpoint
api_name:
- PowerPoint.Font
ms.assetid: ad62daaa-01a5-36cc-5451-e0da0134ac95
ms.date: 06/08/2017
localization_priority: Normal
---


# Font object (PowerPoint)

Represents character formatting for text or a bullet. The **Font** object is a member of the **[Fonts](PowerPoint.Fonts.md)** collection. The **Fonts** collection contains all the fonts used in a presentation.


## Example

The following examples describes how to do the following:


- Return the  **Font** object that represents the font attributes of a specified bullet, a specified range of text, or all text at a specified outline level
    
- Return a **Font** object from the collection of all the fonts used in the presentation
    
Use the [Font](PowerPoint.TextRange.Font.md)property to return the  **Font** object that represents the font attributes for a specific bullet, text range, or outline level. The following example sets the title text on slide one and sets the font properties.




```vb
With ActivePresentation.Slides(1).Shapes.Title _

        .TextFrame.TextRange

    .Text = "Volcano Coffee"

    With .Font

        .Italic = True

        .Name = "Palatino"

        .Color.RGB = RGB(0, 0, 255)

    End With

End With
```

Use  **Fonts** (_index_), where _index_ is the font's name or index number, to return a single **Font** object. The following example checks to see whether font one in the active presentation is embedded in the presentation.




```vb
If ActivePresentation.Fonts(1).Embedded = _

    True Then MsgBox "Font 1 is embedded"
```


## Properties



|Name|
|:-----|
|[Application](PowerPoint.Font.Application.md)|
|[AutoRotateNumbers](PowerPoint.Font.AutoRotateNumbers.md)|
|[BaselineOffset](PowerPoint.Font.BaselineOffset.md)|
|[Bold](PowerPoint.Font.Bold.md)|
|[Color](PowerPoint.Font.Color.md)|
|[Embeddable](PowerPoint.Font.Embeddable.md)|
|[Embedded](PowerPoint.Font.Embedded.md)|
|[Emboss](PowerPoint.Font.Emboss.md)|
|[Italic](PowerPoint.font.italic.md)|
|[Name](PowerPoint.Font.Name.md)|
|[NameAscii](PowerPoint.Font.NameAscii.md)|
|[NameComplexScript](PowerPoint.Font.NameComplexScript.md)|
|[NameFarEast](PowerPoint.Font.NameFarEast.md)|
|[NameOther](PowerPoint.Font.NameOther.md)|
|[Parent](PowerPoint.Font.Parent.md)|
|[Shadow](PowerPoint.Font.Shadow.md)|
|[Size](PowerPoint.Font.Size.md)|
|[Subscript](PowerPoint.Font.Subscript.md)|
|[Superscript](PowerPoint.Font.Superscript.md)|
|[Underline](PowerPoint.Font.Underline.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]