---
title: Story object (Publisher)
keywords: vbapb10.chm5898239
f1_keywords:
- vbapb10.chm5898239
ms.prod: publisher
api_name:
- Publisher.Story
ms.assetid: 0385b4be-0046-9198-a186-0d992601780e
ms.date: 06/01/2019
localization_priority: Normal
---


# Story object (Publisher)

Represents the text in an unlinked text frame, text flowing between linked text frames, or text in a table cell. The **Story** object is a member of the **[TextFrame](publisher.textframe.md)** and **[TextRange](publisher.textrange.md)** objects and the **[Stories](Publisher.Stories.md)** collection.

## Remarks

Use the **[Story](publisher.textframe.story.md)** property of the **TextFrame** or **TextRange** object to return the **Story** object in a text frame or text range. 

Use **[Stories](publisher.document.stories.md)** (_index_), where _index_ is the number of the story, to return an individual **Story** object. 


## Example

This example returns the story in the selected text range, and if it is in a text frame, inserts text into the text range.

```vb
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then .TextRange _ 
 .InsertAfter NewText:=vbLf & "This is a test." 
 End With 
End Sub
```

<br/>

This example determines if the first story in the active publication has a text frame, and if it does, formats the paragraphs in the story with a half inch first line indent and a six-point spacing before each paragraph.

```vb
Sub StoryParagraphFirstLineIndent() 
 With ActiveDocument.Stories(1) 
 If .HasTextFrame Then 
 With .TextFrame.TextRange.ParagraphFormat 
 .FirstLineIndent = InchesToPoints(0.5) 
 .SpaceBefore = 6 
 End With 
 End If 
 End With 
End Sub
```


## Properties

- [Application](Publisher.Story.Application.md)
- [HasTable](Publisher.Story.HasTable.md)
- [HasTextFrame](Publisher.Story.HasTextFrame.md)
- [Parent](Publisher.Story.Parent.md)
- [Table](Publisher.Story.Table.md)
- [TextFrame](Publisher.Story.TextFrame.md)
- [TextRange](Publisher.Story.TextRange.md)
- [Type](Publisher.Story.Type.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]