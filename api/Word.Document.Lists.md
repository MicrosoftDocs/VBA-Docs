---
title: Document.Lists property (Word)
keywords: vbawd10.chm158007360
f1_keywords:
- vbawd10.chm158007360
ms.prod: word
api_name:
- Word.Document.Lists
ms.assetid: 06d5539e-f0a2-0c93-4ade-26403eb6433e
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Lists property (Word)

Returns a  **[Lists](Word.lists.md)** collection that contains all the formatted lists in the specified document. Read-only.


## Syntax

_expression_. `Lists`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example formats the selection as a numbered list. The example then displays a message box that reports the number of lists in the active document.


```vb
Selection.Range.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(2) 
MsgBox "This document has " & ActiveDocument.Lists.Count _ 
 & " lists."
```

This example formats the third list in the active document with the default bulleted list format. If the list is already formatted with a bulleted list format, the example removes the formatting.




```vb
If ActiveDocument.Lists.Count >= 3 Then 
 ActiveDocument.Lists(3).Range.ListFormat.ApplyBulletDefault 
End If
```

This example displays a message box that reports the number of items in each list in MyLetter.doc.




```vb
Set myDoc = Documents("MyLetter.doc") 
i = myDoc.Lists.Count 
For each li in myDoc.Lists 
 Msgbox "List " & i & " has " & li.CountNumberedItems _ 
 & " items." 
 i = i - 1 
Next li
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]