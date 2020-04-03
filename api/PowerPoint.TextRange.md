---
title: TextRange object (PowerPoint)
keywords: vbapp10.chm569000
f1_keywords:
- vbapp10.chm569000
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange
ms.assetid: 7c234107-c423-7ec9-e8bd-a82cc3b345de
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange object (PowerPoint)

Contains the text that's attached to a shape, and properties and methods for manipulating the text.


## Remarks

The following examples describe how to:


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Find and replace text in a text range.
    
- Insert text, the date and time, or the slide number into a text range.
    
- Position the cursor wherever you want in a text range.
    

## Example

Use the [TextRange](PowerPoint.TextFrame.TextRange.md)property of the  **[TextFrame](PowerPoint.TextFrame.md)** object to return a **TextRange** object for any shape you specify. Use the [Text](PowerPoint.TextRange.Text.md)property to return the string of text in the  **TextRange** object. The following example adds a rectangle to _myDocument_ and sets the text it contains.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140) _

    .TextFrame.TextRange.Text = "Here is some test text"
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Text = "Here is some test text"

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange = "Here is some test text"
```

Use the [HasTextFrame](PowerPoint.Shape.HasTextFrame.md)property to determine whether a shape has a text frame, and use the [HasText](PowerPoint.TextFrame.HasText.md)property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```vb
ActiveWindow.Selection.TextRange.Copy
```

Use one of the following methods to return a portion of the text of a **TextRange** object: **[Characters](PowerPoint.TextRange.Characters.md)**, **[Lines](PowerPoint.TextRange.Lines.md)**, **[Paragraphs](PowerPoint.TextRange.Paragraphs.md)**, **[Runs](PowerPoint.TextRange.Runs.md)**, **[Sentences](PowerPoint.TextRange.Sentences.md)**, or **[Words](PowerPoint.TextRange.Words.md)**.

Use the [Find](PowerPoint.TextRange.Find.md)and [Replace](PowerPoint.TextRange.Replace.md)methods to find and replace text in a text range.

Use one of the following methods to insert characters into a **TextRange** object:[InsertAfter](PowerPoint.TextRange.InsertAfter.md), [InsertBefore](PowerPoint.TextRange.InsertBefore.md), [InsertDateTime](PowerPoint.TextRange.InsertDateTime.md), [InsertSlideNumber](PowerPoint.TextRange.InsertSlideNumber.md), or [InsertSymbol](PowerPoint.TextRange.InsertSymbol.md).


## Methods



|Name|
|:-----|
|[AddPeriods](PowerPoint.TextRange.AddPeriods.md)|
|[ChangeCase](PowerPoint.TextRange.ChangeCase.md)|
|[Characters](PowerPoint.TextRange.Characters.md)|
|[Copy](PowerPoint.TextRange.Copy.md)|
|[Cut](PowerPoint.TextRange.Cut.md)|
|[Delete](PowerPoint.TextRange.Delete.md)|
|[Find](PowerPoint.TextRange.Find.md)|
|[InsertAfter](PowerPoint.TextRange.InsertAfter.md)|
|[InsertBefore](PowerPoint.TextRange.InsertBefore.md)|
|[InsertDateTime](PowerPoint.TextRange.InsertDateTime.md)|
|[InsertSlideNumber](PowerPoint.TextRange.InsertSlideNumber.md)|
|[InsertSymbol](PowerPoint.TextRange.InsertSymbol.md)|
|[Lines](PowerPoint.TextRange.Lines.md)|
|[LtrRun](PowerPoint.TextRange.LtrRun.md)|
|[Paragraphs](PowerPoint.TextRange.Paragraphs.md)|
|[Paste](PowerPoint.TextRange.Paste.md)|
|[PasteSpecial](PowerPoint.TextRange.PasteSpecial.md)|
|[RemovePeriods](PowerPoint.TextRange.RemovePeriods.md)|
|[Replace](PowerPoint.TextRange.Replace.md)|
|[RotatedBounds](PowerPoint.TextRange.RotatedBounds.md)|
|[RtlRun](PowerPoint.TextRange.RtlRun.md)|
|[Runs](PowerPoint.TextRange.Runs.md)|
|[Select](PowerPoint.TextRange.Select.md)|
|[Sentences](PowerPoint.TextRange.Sentences.md)|
|[TrimText](PowerPoint.TextRange.TrimText.md)|
|[Words](PowerPoint.TextRange.Words.md)|

## Properties



|Name|
|:-----|
|[ActionSettings](PowerPoint.TextRange.ActionSettings.md)|
|[Application](PowerPoint.TextRange.Application.md)|
|[BoundHeight](PowerPoint.TextRange.BoundHeight.md)|
|[BoundLeft](PowerPoint.TextRange.BoundLeft.md)|
|[BoundTop](PowerPoint.TextRange.BoundTop.md)|
|[BoundWidth](PowerPoint.TextRange.BoundWidth.md)|
|[Count](PowerPoint.TextRange.Count.md)|
|[Font](PowerPoint.TextRange.Font.md)|
|[IndentLevel](PowerPoint.TextRange.IndentLevel.md)|
|[LanguageID](PowerPoint.TextRange.LanguageID.md)|
|[Length](PowerPoint.TextRange.Length.md)|
|[ParagraphFormat](PowerPoint.TextRange.ParagraphFormat.md)|
|[Parent](PowerPoint.TextRange.Parent.md)|
|[Start](PowerPoint.TextRange.Start.md)|
|[Text](PowerPoint.TextRange.Text.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
