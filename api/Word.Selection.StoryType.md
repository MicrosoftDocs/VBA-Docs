---
title: Selection.StoryType property (Word)
keywords: vbawd10.chm158662663
f1_keywords:
- vbawd10.chm158662663
ms.prod: word
api_name:
- Word.Selection.StoryType
ms.assetid: 17709b74-ea6b-9d58-885d-01e6b2ddac55
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.StoryType property (Word)

Returns the story type for the specified selection. Read-only  **WdStoryType**.


## Syntax

_expression_. `StoryType`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example returns the story type of the selection.


```vb
story = Selection.StoryType
```

This example closes the footnote pane if the selection is contained in the footnote story.




```vb
ActiveDocument.ActiveWindow.View.Type = wdNormalView 
If Selection.StoryType = wdFootnotesStory Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]