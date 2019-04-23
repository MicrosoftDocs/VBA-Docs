---
title: Bookmarks.Item method (Word)
ms.prod: word
api_name:
- Word.Bookmarks.Item
ms.assetid: 95650b7b-fe74-09a4-60a6-a716407e8a34
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmarks.Item method (Word)

Returns an individual  **Bookmark** object in a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ Required. A variable that represents a '[Bookmarks](Word.bookmarks.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The individual object to be returned. Can be a  **Long** indicating the ordinal position or a **String** representing the name of the individual object.|

## Return value

Bookmark


## Example

This example selects the bookmark named "temp" in the active document.


```vb
Sub BookmarkItem() 
 If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 ActiveDocument.Bookmarks.Item("temp").Select 
 End If 
End Sub
```


## See also


[Bookmarks Collection Object](Word.bookmarks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]