---
title: Application.Documents property (Word)
keywords: vbawd10.chm158334982
f1_keywords:
- vbawd10.chm158334982
ms.prod: word
api_name:
- Word.Application.Documents
ms.assetid: 7e477cb3-ae65-685a-0083-1826efe86703
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Documents property (Word)

Returns a **[Documents](Word.documents.md)** collection that represents all the open documents. Read-only.


## Syntax

_expression_.**Documents**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


> [!NOTE] 
> A document displayed in a Protected View window is not a member of the **[Documents](Word.Application.Documents.md)** collection. Instead, use the [Document](Word.Document.md) property of the [ProtectedViewWindow](Word.ProtectedViewWindow.md) object to access a document that is displayed in a Protected View window.


## Example

This example creates a new document based on the Normal template and then displays the **Save As** dialog box.


```vb
Documents.Add.Save
```

This example saves open documents that have changed since they were last saved.




```vb
Dim docLoop As Document 
 
For Each docLoop In Documents 
   If docLoop.Saved = False Then docLoop.Save 
Next docLoop
```

This example prints each open document after setting the left and right margins to 0.5 inch.




```vb
Dim docLoop As Document 
 
For Each docLoop In Documents 
    With docLoop 
        .PageSetup.LeftMargin = InchesToPoints(0.5) 
        .PageSetup.RightMargin = InchesToPoints(0.5) 
        .PrintOut 
    End With 
Next docLoop
```

This example opens Doc.doc as a read-only document.




```vb
Documents.Open FileName:="C:\Files\Doc.doc", ReadOnly:=True
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
