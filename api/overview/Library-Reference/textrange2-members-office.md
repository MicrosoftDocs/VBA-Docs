---
title: TextRange2 members (Office)
description: Represents the text frame in a Shape or ShapeRange object.
ms.prod: office
ms.assetid: 26daffff-b9ef-fd94-f5b7-ed3a09840cb6
ms.date: 01/30/2019
localization_priority: Normal
---


# TextRange2 members (Office)

Represents the text frame in a **Shape** or **ShapeRange** object.

## Methods

|Name|Description|
|:-----|:-----|
|[AddPeriods](../../Office.TextRange2.AddPeriods.md)|Adds period (.) punctuation to the right side of the text contained in **TextRange2** object for left-to-right languages and on the left side for right-to-left languages.|
|[ChangeCase](../../Office.TextRange2.ChangeCase.md)|Changes the case of a **TextRange2** object to one of the values in the **MsoTextChangeCase** enumeration.|
|[Copy](../../Office.TextRange2.Copy.md)|Copies a **TextRange2** object.|
|[Cut](../../Office.TextRange2.Cut.md)|Removes a portion or all of the text from a range of text.|
|[Delete](../../Office.TextRange2.Delete.md)|Deletes a **TextRange2** object.|
|[Find](../../Office.TextRange2.Find.md)|Searches a **TextRange2** object for a subset of text.|
|[InsertAfter](../../Office.TextRange2.InsertAfter.md)|Inserts text to the right of the existing text in the **TextRange2** object.|
|[InsertBefore](../../Office.TextRange2.InsertBefore.md)|Inserts text to the left of the existing text in the **TextRange2** object.|
|[InsertChartField](../../Office.TextRange2.InsertChartField.md)|Inserts a field into the body of a data label in a chart. |
|[InsertSymbol](../../Office.TextRange2.InsertSymbol.md)|Inserts a symbol from the specified font set into the range of text represented by the **TextRange2** object.|
|[Item](../../Office.TextRange2.Item.md)|Gets the range of text specified by the index number from the **TextRange2** object.|
|[LtrRun](../../Office.TextRange2.LtrRun.md)|Returns a **TextRange2** object that represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.|
|[Paste](../../Office.TextRange2.Paste.md)|Pastes the contents of the Clipboard into the **TextRange2** object.|
|[PasteSpecial](../../Office.TextRange2.PasteSpecial.md)|Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a **TextRange2** object including the text range that was pasted.|
|[RemovePeriods](../../Office.TextRange2.RemovePeriods.md)|Removes all period (.) punctuation from the text in the **TextRange2** object.|
|[Replace](../../Office.TextRange2.Replace.md)|Finds specific text in a text range, replaces the found text with a specified string, and returns a **TextRange2** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.|
|[RotatedBounds](../../Office.TextRange2.RotatedBounds.md)|Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.|
|[RtlRun](../../Office.TextRange2.RtlRun.md)|Returns a **TextRange2** object that represents the specified subset of right-to-left text runs. A text run consists of a range of characters that share the same font attributes.|
|[Select](../../Office.TextRange2.Select.md)|Selects the **TextRange2** object.|
|[TrimText](../../Office.TextRange2.TrimText.md)|Returns a **TextRange2** object that represents the specified text that has the whitespace removed.|

## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.TextRange2.Application.md)|When used without an object qualifier, this property returns an **Application** object that represents the current instance of the Microsoft Office application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the **TextRange2** object. When used with an OLE **Automation** object, it returns the object's application. Read-only.|
|[BoundHeight](../../Office.TextRange2.BoundHeight.md)|Gets the height, in points, of the text bounding box for the specified text. Read-only.|
|[BoundLeft](../../Office.TextRange2.BoundLeft.md)|Gets the left coordinate, in points, of the text bounding box for the specified text. Read-only.|
|[BoundTop](../../Office.TextRange2.BoundTop.md)|Gets the top coordinate, in points, of the text bounding box for the specified text. Read-only.|
|[BoundWidth](../../Office.TextRange2.BoundWidth.md)|Gets the width, in points, of the text bounding box for the specified text. Read-only.|
|[Characters](../../Office.TextRange2.Characters.md)|Read-only.|
|[Count](../../Office.TextRange2.Count.md)|Gets a **Long** indicating the number of items in the **TextRange2** collection. Read-only.|
|[Creator](../../Office.TextRange2.Creator.md)|Gets a 32-bit integer that indicates the application in which the **TextRange2** object was created. Read-only.|
|[Font](../../Office.TextRange2.Font.md)|Returns a **Font** object that represents character formatting for the **TextRange2** object. Read-only.|
|[LanguageID](../../Office.TextRange2.LanguageID.md)|Gets or sets the **MsoLanguageID** value of the **TextRange2** object. Read/write.|
|[Length](../../Office.TextRange2.Length.md)|Get a **Long** that represents the length of a text range. Read-only.|
|[Lines](../../Office.TextRange2.Lines.md)|Returns a **TextRange2** object that represents the specified subset of text lines. Read-only.|
|[MathZones](../../Office.TextRange2.MathZones.md)|Sets the starting point and length of a math zone within a text range. Read-only.|
|[ParagraphFormat](../../Office.TextRange2.ParagraphFormat.md)|Returns a **ParagraphFormat** object that represents paragraph formatting for the specified text. Read-only.|
|[Paragraphs](../../Office.TextRange2.Paragraphs.md)|Gets a **TextRange2** object that represents the specified subset of text paragraphs. Read-only.|
|[Parent](../../Office.TextRange2.Parent.md)|Gets the **Parent** object for the **TextRange2** object. Read-only.|
|[Runs](../../Office.TextRange2.Runs.md)|Gets a **TextRange2** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes. Read-only.|
|[Sentences](../../Office.TextRange2.Sentences.md)|Returns a **TextRange2** object that represents the specified subset of text sentences. Read-only.|
|[Start](../../Office.TextRange2.Start.md)|Gets a **Long** value indicating the starting point of the specified text range. Read-only.|
|[Text](../../Office.TextRange2.Text.md)|Gets or sets a **String** value that represents the text in a text range. Read/write.|
|[Words](../../Office.TextRange2.Words.md)|Gets a **TextRange2** object that represents the specified subset of text words. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]