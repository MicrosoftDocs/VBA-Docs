---
title: Bookmark object (Word)
keywords: vbawd10.chm2408
f1_keywords:
- vbawd10.chm2408
ms.prod: word
api_name:
- Word.Bookmark
ms.assetid: be6b0c7b-60ca-97e7-ef19-6de335da3197
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmark object (Word)

Represents a single bookmark in a document, selection, or range. The  **Bookmark** object is a member of the **[Bookmarks](Word.bookmarks.md)** collection. The **Bookmarks** collection includes all the bookmarks listed in the **Bookmark** dialog box (**Insert** menu).


## Remarks

Using the Bookmark Object

Use  **Bookmarks** (_index_), where _index_ is the bookmark name or index number, to return a single **Bookmark** object. You must exactly match the spelling (but not necessarily the capitalization) of the bookmark name. The following example selects the bookmark named "temp" in the active document.




```vb
ActiveDocument.Bookmarks("temp").Select
```

The index number represents the position of the bookmark in the  **[Selection](Word.Selection.md)** or **[Range](Word.Range.md)** object. For the **[Document](Word.Document.md)** object, the index number represents the position of the bookmark in the alphabetical list of bookmarks in the **Bookmarks** dialog box (click **Name** to sort the list of bookmarks alphabetically). The following example displays the name of the second bookmark in the **Bookmarks** collection.




```vb
MsgBox ActiveDocument.Bookmarks(2).Name
```

Use the  **[Add](Word.Bookmarks.Add.md)** method to add a bookmark to a document range. The following example marks the selection by adding a bookmark named "temp."




```vb
ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
```

Remarks

Use the  **BookmarkID** property with a range or selection object to return the index number of a **Bookmark** object in the **Bookmarks** collection. The following example displays the index number of the bookmark named "temp" in the active document.




```vb
MsgBox ActiveDocument.Bookmarks("temp").Range.BookmarkID
```

You can use [predefined bookmarks](../word/Concepts/Miscellaneous/predefined-bookmarks.md)with the  **Bookmarks** property. The following example sets the bookmark named "currpara" to the location marked by the predefined bookmark named "\Para".




```vb
ActiveDocument.Bookmarks("\Para").Copy "currpara"
```

Use the  **[Exists](Word.Bookmarks.Exists.md)** method to determine whether a bookmark already exists in the selection, range, or document. The following example ensures that the bookmark named "temp" exists in the active document before selecting the bookmark.




```vb
If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 ActiveDocument.Bookmarks("temp").Select 
End If
```


## Methods

- [Copy](Word.Bookmark.Copy.md)
- [Delete](Word.Bookmark.Delete.md)
- [Select](Word.Bookmark.Select.md)

## Properties

- [Application](Word.Bookmark.Application.md)
- [Column](Word.Bookmark.Column.md)
- [Creator](Word.Bookmark.Creator.md)
- [Empty](Word.Bookmark.Empty.md)
- [End](Word.Bookmark.End.md)
- [Name](Word.Bookmark.Name.md)
- [Parent](Word.Bookmark.Parent.md)
- [Range](Word.Bookmark.Range.md)
- [Start](Word.Bookmark.Start.md)
- [StoryType](Word.Bookmark.StoryType.md)


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
