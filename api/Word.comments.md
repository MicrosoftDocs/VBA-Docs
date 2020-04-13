---
title: Comments object (Word)
ms.prod: word
ms.assetid: e384b37a-50e3-a214-52a8-6fda2acc4991
ms.date: 06/08/2017
localization_priority: Normal
---


# Comments object (Word)

A collection of  **[Comment](Word.Comment.md)** objects that represent the comments in a selection, range, or document.


## Remarks

Use the **Comments** property to return the **Comments** collection. The following example displays comments made by Don Funk in the active document.


```vb
ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments 
ActiveDocument.Comments.ShowBy = "Don Funk"
```

Use the **[Add](Word.Comments.Add.md)** method to add a comment at the specified range. The following example adds a comment immediately after the selection.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Comments.Add Range:=Selection.Range, _ 
 Text:="review this"
```

Use  **Comments** (Index), where Index is the index number, to return a single **Comment** object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document.




```vb
MsgBox ActiveDocument.Comments(1).Author
```

The following example displays the initials of the author of the first comment in the selection.




```vb
If Selection.Comments.Count >= 1 Then MsgBox _ 
 Selection.Comments(1).Initial
```


## Methods



|Name|
|:-----|
|[Add](Word.Comments.Add.md)|
|[Item](Word.Comments.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Comments.Application.md)|
|[Count](Word.Comments.Count.md)|
|[Creator](Word.Comments.Creator.md)|
|[Parent](Word.Comments.Parent.md)|
|[ShowBy](Word.Comments.ShowBy.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]