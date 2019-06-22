---
title: Global.Documents property (Word)
keywords: vbawd10.chm163119105
f1_keywords:
- vbawd10.chm163119105
ms.prod: word
api_name:
- Word.Global.Documents
ms.assetid: a86bad22-aabf-dd0d-4b23-fc608d5db4c1
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Documents property (Word)

Returns a  **[Documents](Word.documents.md)** collection that represents all the open documents. Read-only.


## Syntax

_expression_.**Documents**

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a new document based on the Normal template and then displays the Save As dialog box.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]