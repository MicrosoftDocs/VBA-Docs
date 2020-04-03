---
title: Working with Document Objects
ms.prod: word
ms.assetid: af304f65-6cdd-ff7d-a81f-cce0161f2b47
ms.date: 06/08/2019
localization_priority: Normal
---

# Working with Document Objects

In Visual Basic, the methods for modifying files are methods of the **[Document](../../../api/Word.Document.md)** object or the **[Documents](../../../api/Word.documents.md)** collection. This topic includes Visual Basic examples related to the tasks identified in the following sections.

## Creating a new document

The **Documents** collection includes all of the open documents. To create a new document, use the **[Add](../../../api/Word.Documents.Add.md)** method to add a **Document** object to the **Documents** collection. The following instruction creates a document.


```vb
Documents.Add
```

A better way to create a document is to assign the return value to an object variable. The **Add** method returns a **Document** object that refers to the new document. In the following example, the **Document** object returned by the **Add** method is assigned to an object variable. Then, several properties and methods of the **Document** object are set. You can easily control the new document using an object variable.




```vb
Sub NewSampleDoc() 
    Dim docNew As Document 
    Set docNew = Documents.Add 
    With docNew 
        .Content.Font.Name = "Tahoma" 
        .SaveAs FileName:="Sample.doc" 
    End With 
End Sub
```


## Opening a document

To open an existing document, use the **[Open](../../../api/Word.Documents.Open.md)** method with the **Documents** collection. The following instruction opens a document named Sample.doc located in the MyFolder folder.


```vb
Sub OpenDocument() 
    Documents.Open FileName:="C:\MyFolder\Sample.doc" 
End Sub
```


## Saving an existing document

To save a single document, use the **[Save](../../../api/Word.Document.Save.md)** method with the **Document** object. The following instruction saves the document named Sales.doc.


```vb
Sub SaveDocument() 
    Documents("Sales.doc").Save 
End Sub
```

You can save all open documents by applying the **Save** method to the **Documents** collection. The following instruction saves all open documents.




```vb
Sub SaveAllOpenDocuments() 
    Documents.Save 
End Sub
```


## Saving a new document

To save a single document, use the **[SaveAs2](../../../api/Word.SaveAs2.md)** method with a **Document** object. The following instruction saves the active document as "Temp.doc" in the current folder.


```vb
Sub SaveNewDocument() 
    ActiveDocument.SaveAs FileName:="Temp.doc"
End Sub
```

The **_FileName_** argument can include only the file name or the complete path (for example, "C:\Documents\Temporary File.doc").


## Closing documents

To close a single document, use the **[Close](../../../api/Word.Document.Close(method).md)** method with a **Document** object. The following instruction closes and saves the document named Sales.doc.

```vb
Sub CloseDocument()
    Documents("Sales.doc").Close SaveChanges:=wdSaveChanges 
End Sub
```

You can close all open documents by applying the **[Close](../../../api/Word.Documents.Close.md)** method of the **Documents** collection. The following instruction closes all documents without saving changes.

```vb
Sub CloseAllDocuments() 
    Documents.Close SaveChanges:=wdDoNotSaveChanges
End Sub
```

The following example prompts the user to save each document before the document is closed.

```vb
Sub PromptToSaveAndClose()
    Dim doc As Document
    For Each doc In Documents 
        doc.Close SaveChanges:=wdPromptToSaveChanges
    Next
End Sub
```


## Activating a document

To change the active document, use the **[Activate](../../../api/Word.Document.Activate.md)** method with a **Document** object. The following instruction activates the open document named Sales.doc.


```vb
Sub ActivateDocument() 
    Documents("Sales.doc").Activate 
End Sub
```


## Determining if a document is open

To determine if a document is open, you can enumerate the **Documents** collection by using a **For Each...Next** statement. The following example activates the document named Sample.doc if the document is open, or opens Sample.doc if it is not currently open.


```vb
Sub ActivateOrOpenDocument() 
    Dim doc As Document 
    Dim docFound As Boolean 
 
    For Each doc In Documents 
        If InStr(1, doc.Name, "sample.doc", 1) Then 
            doc.Activate 
            docFound = True 
            Exit For 
        Else 
            docFound = False 
        End If 
    Next doc 
 
    If docFound = False Then Documents.Open FileName:="Sample.doc" 
End Sub
```


## Referring to the active document

Instead of referring to a document by name or by index number—for example,  `Documents("Sales.doc")` —the **[ActiveDocument](../../../api/Word.Application.ActiveDocument.md)** property returns a **Document** object that refers to the active document (the document with the focus). The following example displays the name of the active document, or if there are no documents open, it displays a message.


```vb
Sub ActiveDocumentName() 
    If Documents.Count >= 1 Then 
        MsgBox ActiveDocument.Name 
    Else 
        MsgBox "No documents are open" 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
