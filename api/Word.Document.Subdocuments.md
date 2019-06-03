---
title: Document.Subdocuments property (Word)
keywords: vbawd10.chm158007341
f1_keywords:
- vbawd10.chm158007341
ms.prod: word
api_name:
- Word.Document.Subdocuments
ms.assetid: 4d0047da-03ef-67da-61ed-8bdbeaa55024
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Subdocuments property (Word)

Returns a  **[Subdocuments](Word.subdocuments.md)** collection that represents all the subdocuments in the specified document. Read-only.


## Syntax

_expression_. `Subdocuments`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of subdocuments embedded in the active document.


```vb
MsgBox ActiveDocument.Subdocuments.Count
```

This example displays the path and file name of each subdocument in the active document.




```vb
For Each subdoc In ActiveDocument.Subdocuments 
 If subdoc.HasFile = True Then 
 MsgBox subdoc.Path & Application.PathSeparator _ 
 & subdoc.Name 
 Else 
 MsgBox "This subdocument has not been saved." 
 End If 
Next subdoc
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]