---
title: Document.Hyperlinks property (Word)
keywords: vbawd10.chm158007357
f1_keywords:
- vbawd10.chm158007357
ms.prod: word
api_name:
- Word.Document.Hyperlinks
ms.assetid: b8db5b89-0a2a-ffe9-c353-1fa77190af75
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Hyperlinks property (Word)

Returns a  **[Hyperlinks](Word.hyperlinks.md)** collection that represents all the hyperlinks in the specified document. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the target address of the second hyperlink in Home.doc.


```vb
If Documents("Home.doc").Hyperlinks.Count >= 2 Then 
 MsgBox Documents("Home.doc").Hyperlinks(2).Name 
End If
```

This example displays the name of every hyperlink in the active document that includes the word "Microsoft" in its address.




```vb
For Each aHyperlink In ActiveDocument.Hyperlinks 
 If InStr(LCase(aHyperlink.Address), "microsoft") <> 0 Then 
 MsgBox aHyperlink.Name 
 End If 
Next aHyperlink
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]