---
title: Characters Object (Word)
ms.prod: word
ms.assetid: 6d22ae7a-128d-134d-9136-1cdd5a8d9941
ms.date: 06/08/2017
---


# Characters Object (Word)

A collection of characters in a selection, range, or document. There is no Character object; instead, each item in the  **Characters** collection is a **[Range](Word.Range.md)** object that represents one character.


## Remarks

Use the  **Characters** property of a **[Document](Word.Document.md)**, **Range**, or **[Selection](Word.Selection.md)** object to return the **Characters** collection. The following example displays how many characters are selected.


```
MsgBox Selection.Characters.Count &amp; " characters are selected"
```

Use  **Characters** (Index), where Index is the index number, to return a **Range** object that represents one character. The index number represents the position of a character in the **Characters** collection. The following example formats the first letter in the selection as 24-point bold.




```
With Selection.Characters(1) 
 .Bold = True 
 .Font.Size = 24 
End With
```

Remarks

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

There is no  **Add** method for the **Characters** collection. Instead, use the **InsertAfter** or **InsertBefore** method to add characters to a **Range** object. The following example inserts a new paragraph after the first paragraph in the active document.




```
With ActiveDocument 
 .Paragraphs(1).Range.InsertParagraphAfter 
 .Paragraphs(2).Range.InsertBefore "New Text" 
End With
```


## Methods



|**Name**|
|:-----|
|[Item](Word.Characters.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Word.Characters.Application.md)|
|[Count](Word.Characters.Count.md)|
|[Creator](Word.Characters.Creator.md)|
|[First](Word.Characters.First.md)|
|[Last](Word.Characters.Last.md)|
|[Parent](Word.Characters.Parent.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
