---
title: Bookmarks object (Word)
ms.prod: word
ms.assetid: 827bed64-3034-0eb4-401d-f117cdb98898
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmarks object (Word)

A collection of  **[Bookmark](Word.Bookmark.md)** objects that represent the bookmarks in the specified selection, range, or document.


## Remarks

Use the **Bookmarks** property to return the **Bookmarks** collection for a document, range, or selection. The following example ensures that the bookmark named "temp" exists in the active document before selecting the bookmark.


```vb
If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 ActiveDocument.Bookmarks("temp").Select 
End If
```

Use the **[Add](Word.Bookmarks.Add.md)** method to set a bookmark for a range in a document. The following example marks the selection by adding a bookmark named "temp".




```vb
ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
```

Use  **Bookmarks** (_index_), where _index_ is the bookmark name or index number, to return a single **Bookmark** object. You must exactly match the spelling (but not necessarily the capitalization) of the bookmark name. The following example selects the bookmark named "temp" in the active document.




```vb
ActiveDocument.Bookmarks("temp").Select
```

The index number represents the position of the bookmark in the **[Selection](Word.Selection.md)** or **[Range](Word.Range.md)** object. For the **[Document](Word.Document.md)** object, the index number represents the position of the bookmark in the alphabetical list of bookmarks in the **Bookmarks** dialog box (click **Name** to sort the list of bookmarks alphabetically). The following example displays the name of the second bookmark in the **Bookmarks** collection.




```vb
MsgBox ActiveDocument.Bookmarks(2).Name
```

Remarks

The **[ShowHidden](Word.Bookmarks.ShowHidden.md)** property effects the number of elements in the **Bookmarks** collection. If **ShowHidden** is **True**, hidden bookmarks are included in the **Bookmarks** collection.

## Methods

- [Add](Word.Bookmarks.Add.md)
- [Exists](Word.Bookmarks.Exists.md)
- [Item](Word.Bookmarks.Item.md)

## Properties

- [Application](Word.Bookmarks.Application.md)
- [Count](Word.Bookmarks.Count.md)
- [Creator](Word.Bookmarks.Creator.md)
- [DefaultSorting](Word.Bookmarks.DefaultSorting.md)
- [Parent](Word.Bookmarks.Parent.md)
- [ShowHidden](Word.Bookmarks.ShowHidden.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
