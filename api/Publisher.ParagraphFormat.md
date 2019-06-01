---
title: ParagraphFormat object (Publisher)
keywords: vbapb10.chm5505023
f1_keywords:
- vbapb10.chm5505023
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat
ms.assetid: 0e5b1c20-564e-ef5c-f24d-1143dcaadcd8
ms.date: 06/01/2019
localization_priority: Normal
---


# ParagraphFormat object (Publisher)

Represents all the formatting for a paragraph.

## Remarks

Use the **[TextStyle.ParagraphFormat](Publisher.TextStyle.ParagraphFormat.md)** property to return the **ParagraphFormat** object for a paragraph or paragraphs. The **ParagraphFormat** property returns the **ParagraphFormat** object for a selection, range, or style. 

Use the **Duplicate** method to copy an existing **ParagraphFormat** object. 

## Example

The following example centers the paragraph at the cursor position. This example assumes that the first shape is a text box and not another type of shape.

```vb
Sub CenterParagraph() 
 Selection.TextRange.ParagraphFormat _ 
 .Alignment = pbParagraphAlignmentCenter 
End Sub
```

<br/>

The following example duplicates the paragraph formatting of the first paragraph in the active publication and stores the formatting in a variable. This example duplicates an existing **ParagraphFormat** object and then changes the left indent to one inch, creates a new textbox, inserts text into it, and applies the paragraph formatting of the duplicated paragraph format to the text.

```vb
Sub DuplicateParagraphFormating() 
 Dim pfmtDup As ParagraphFormat 
 
 Set pfmtDup = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 
 pfmtDup.LeftIndent = Application.InchesToPoints(1) 
 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddTextbox(pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=200, Height:=100) 
 With .TextFrame.TextRange 
 .Text = "This is a test of how to use " & _ 
 "the ParagraphFormat object." 
 .ParagraphFormat = pfmtDup 
 End With 
 End With 
 End With 
 
End Sub
```


## Methods

- [Duplicate](Publisher.ParagraphFormat.Duplicate.md)
- [Reset](Publisher.ParagraphFormat.Reset.md)
- [SetLineSpacing](Publisher.ParagraphFormat.SetLineSpacing.md)
- [SetListType](Publisher.ParagraphFormat.SetListType.md)

## Properties

- [Alignment](Publisher.ParagraphFormat.Alignment.md)
- [Application](Publisher.ParagraphFormat.Application.md)
- [AttachedToText](Publisher.ParagraphFormat.AttachedToText.md)
- [CharBasedFirstLineIndent](Publisher.ParagraphFormat.CharBasedFirstLineIndent.md)
- [FirstLineIndent](Publisher.ParagraphFormat.FirstLineIndent.md)
- [KashidaPercentage](Publisher.ParagraphFormat.KashidaPercentage.md)
- [KeepLinesTogether](Publisher.ParagraphFormat.KeepLinesTogether.md)
- [KeepWithNext](Publisher.ParagraphFormat.KeepWithNext.md)
- [LeftIndent](Publisher.ParagraphFormat.LeftIndent.md)
- [LineSpacing](Publisher.ParagraphFormat.LineSpacing.md)
- [LineSpacingRule](Publisher.ParagraphFormat.LineSpacingRule.md)
- [ListBulletFontName](Publisher.ParagraphFormat.ListBulletFontName.md)
- [ListBulletFontSize](Publisher.ParagraphFormat.ListBulletFontSize.md)
- [ListBulletText](Publisher.ParagraphFormat.ListBulletText.md)
- [ListIndent](Publisher.ParagraphFormat.ListIndent.md)
- [ListNumberSeparator](Publisher.ParagraphFormat.ListNumberSeparator.md)
- [ListNumberStart](Publisher.ParagraphFormat.ListNumberStart.md)
- [ListType](Publisher.ParagraphFormat.ListType.md)
- [LockToBaseLine](Publisher.ParagraphFormat.LockToBaseLine.md)
- [Parent](Publisher.ParagraphFormat.Parent.md)
- [RightIndent](Publisher.ParagraphFormat.RightIndent.md)
- [SpaceAfter](Publisher.ParagraphFormat.SpaceAfter.md)
- [SpaceBefore](Publisher.ParagraphFormat.SpaceBefore.md)
- [StartInNextTextBox](Publisher.ParagraphFormat.StartInNextTextBox.md)
- [Tabs](Publisher.ParagraphFormat.Tabs.md)
- [TextDirection](Publisher.ParagraphFormat.TextDirection.md)
- [TextStyle](Publisher.ParagraphFormat.TextStyle.md)
- [UseCharBasedFirstLineIndent](Publisher.ParagraphFormat.UseCharBasedFirstLineIndent.md)
- [WidowControl](Publisher.ParagraphFormat.WidowControl.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]