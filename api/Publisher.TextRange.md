---
title: TextRange Object (Publisher)
keywords: vbapb10.chm5373951
f1_keywords:
- vbapb10.chm5373951
ms.prod: publisher
api_name:
- Publisher.TextRange
ms.assetid: 566f240b-d2a6-8cb3-9eb7-68328d6c28bd
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange Object (Publisher)

Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text. This topic describes how to: 


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Insert text, the date and time, or the page number into a text range.
    

## Example

Use the  **[TextRange](./Publisher.TextFrame.TextRange.md)** property of the **[TextFrame](Publisher.TextFrame.md)** object to return a **TextRange** object for any shape you specify. Use the **[Text](./Publisher.TextRange.Text.md)** property to return the string of text in the **TextRange** object. The following example adds a rectangle to the active publication and sets the text it contains.


```vb
Sub AddTextToShape() 
    With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
        Left:=72, Top:=72, Width:=250, Height:=140) 
        .TextFrame.TextRange.Text = "Here is some test text" 
    End With 
End Sub
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange.text = "Here is some test text" 
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange = "Here is some test text"
```

Use the  **[HasTextFrame](./Publisher.ShapeRange.HasTextFrame.md)** property to determine whether a shape has a text frame, and use the **[HasText](./Publisher.TextFrame.HasText.md)** property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```vb
Sub CopyAndPasteText() 
    With ActiveDocument 
        .Selection.TextRange.Copy 
        .Pages(1).Shapes(1).TextFrame.TextRange.Paste 
    End With 
End Sub
```

Use one of the following methods to return a portion of the text of a  **TextRange** object: **[Characters](./Publisher.TextRange.Characters.md)**, **[Lines](./Publisher.TextRange.Lines.md)**, **[Paragraphs](./Publisher.TextRange.Paragraphs.md)**, or **[Words](./Publisher.TextRange.Words.md)**. The following example formats the second word in the first shape on the first page of the active publication. For this example to work, the specified shape must contain text.




```vb
Sub FormatWords() 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange.Words(2).Font 
        .Bold = msoTrue 
        .Size = 15 
        .Name = "Text Name" 
    End With 
End Sub
```

Use one of the following methods to insert characters into a  **TextRange** object: **[InsertAfter](./Publisher.TextRange.InsertAfter.md)**, **[InsertBefore](./Publisher.TextRange.InsertBefore.md)**, **[InsertDateTime](./Publisher.TextRange.InsertDateTime.md)**, **[InsertPageNumber](./Publisher.TextRange.InsertPageNumber.md)**, or **[InsertSymbol](./Publisher.TextRange.InsertSymbol.md)**. This example inserts a new line with text after any existing text in the first shape on the first page of the active publication.




```vb
Sub InsertNewText() 
    Dim intCount As Integer 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange 
        For intCount = 1 To 3 
            .InsertAfter vbLf &amp; "This is a test." 
        Next intCount 
    End With 
End Sub
```


## Methods



|Name|
|:-----|
|[Characters](./Publisher.TextRange.Characters.md)|
|[Collapse](./Publisher.TextRange.Collapse.md)|
|[Copy](./Publisher.TextRange.Copy.md)|
|[Cut](./Publisher.TextRange.Cut.md)|
|[Delete](./Publisher.TextRange.Delete.md)|
|[Expand](./Publisher.TextRange.Expand.md)|
|[InsertAfter](./Publisher.TextRange.InsertAfter.md)|
|[InsertBarcode](./Publisher.TextRange.InsertBarcode.md)|
|[InsertBefore](./Publisher.TextRange.InsertBefore.md)|
|[InsertDateTime](./Publisher.TextRange.InsertDateTime.md)|
|[InsertMailMergeField](./Publisher.TextRange.InsertMailMergeField.md)|
|[InsertPageNumber](./Publisher.TextRange.InsertPageNumber.md)|
|[InsertSymbol](./Publisher.TextRange.InsertSymbol.md)|
|[Lines](./Publisher.TextRange.Lines.md)|
|[Move](./Publisher.TextRange.Move.md)|
|[MoveEnd](./Publisher.TextRange.MoveEnd.md)|
|[MoveStart](./Publisher.TextRange.MoveStart.md)|
|[Paragraphs](./Publisher.TextRange.Paragraphs.md)|
|[Paste](./Publisher.TextRange.Paste.md)|
|[Select](./Publisher.TextRange.Select.md)|
|[Words](./Publisher.TextRange.Words.md)|

## Properties



|Name|
|:-----|
|[Application](./Publisher.TextRange.Application.md)|
|[BoundHeight](./Publisher.TextRange.BoundHeight.md)|
|[BoundLeft](./Publisher.TextRange.BoundLeft.md)|
|[BoundTop](./Publisher.TextRange.BoundTop.md)|
|[BoundWidth](./Publisher.TextRange.BoundWidth.md)|
|[ContainingObject](./Publisher.TextRange.ContainingObject.md)|
|[DropCap](./Publisher.TextRange.DropCap.md)|
|[Duplicate](./Publisher.TextRange.Duplicate.md)|
|[End](./Publisher.TextRange.End.md)|
|[Fields](./Publisher.TextRange.Fields.md)|
|[Find](./Publisher.TextRange.Find.md)|
|[Font](./Publisher.TextRange.Font.md)|
|[Hyperlinks](./Publisher.TextRange.Hyperlinks.md)|
|[InlineShapes](./Publisher.TextRange.InlineShapes.md)|
|[LanguageID](./Publisher.TextRange.LanguageID.md)|
|[Length](./Publisher.TextRange.Length.md)|
|[LinesCount](./Publisher.TextRange.LinesCount.md)|
|[MajorityFont](./Publisher.TextRange.MajorityFont.md)|
|[MajorityParagraphFormat](./Publisher.TextRange.MajorityParagraphFormat.md)|
|[ParagraphFormat](./Publisher.TextRange.ParagraphFormat.md)|
|[ParagraphsCount](./Publisher.TextRange.ParagraphsCount.md)|
|[Parent](./Publisher.TextRange.Parent.md)|
|[Script](./Publisher.TextRange.Script.md)|
|[Start](./Publisher.TextRange.Start.md)|
|[Story](./Publisher.TextRange.Story.md)|
|[Text](./Publisher.TextRange.Text.md)|
|[WordsCount](./Publisher.TextRange.WordsCount.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]