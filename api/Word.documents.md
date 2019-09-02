---
title: Documents object (Word)
ms.prod: word
ms.assetid: fc4ac973-19c1-703a-5538-f4426b8b7564
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents object (Word)

A collection of all the **[Document](Word.Document.md)** objects that are currently open in Word.


## Remarks

Use the **Documents** property to return the **Documents** collection. The following example displays the names of the open documents.


```vb
For Each aDoc In Documents 
 aName = aName & aDoc.Name & vbCr 
Next aDoc 
MsgBox aName
```

Use the **[Add](Word.Documents.Add.md)** method to create a new empty document and add it to the **Documents** collection. The following example creates a new document based on the Normal template.




```vb
Documents.Add
```

Use the **[Open](Word.Documents.Open.md)** method to open a file. The following example opens the document named "Sales.doc."




```vb
Documents.Open FileName:="C:\My Documents\Sales.doc"
```

Use **[Documents](Word.Application.Documents.md)** (Index), where Index is the document name or index number to return a single **Document** object. The following instruction closes the document named "Report.doc" without saving changes.




```vb
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the **Documents** collection. The following example activates the first document in the **Documents** collection.




```vb
Documents(1).Activate
```

The following example enumerates the **Documents** collection to determine whether the document named "Report.doc" is open. If this document is contained in the **Documents** collection, the document is activated; otherwise, it is opened.




```vb
For Each doc In Documents 
 If doc.Name = "Report.doc" Then found = True 
Next doc 
If found <> True Then 
 Documents.Open FileName:="C:\Documents\Report.doc" 
Else 
 Documents("Report.doc").Activate 
End If
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
