---
title: Documents object (Publisher)
keywords: vbapb10.chm8716287
f1_keywords:
- vbapb10.chm8716287
ms.prod: publisher
api_name:
- Publisher.Documents
ms.assetid: 855b1677-4072-1e17-c22c-6db08e0c7569
ms.date: 05/31/2019
localization_priority: Normal
---


# Documents object (Publisher)

Represents all open publications. The **Documents** collection contains all **[Document](Publisher.Document.md)** objects that are open in Microsoft Publisher.

## Remarks

Use the **[Documents](publisher.application.documents.md)** property to return the **Documents** collection. 

Use the **Add** method to add a new document to the collection. A new and visible instance of Publisher is created when the **Add** method is called. 

Use the **Item** (_index_) property, where _index_ is the index number or document name as a **String**, to return a specific document object. 


## Example

The following example lists all the open publications.

```vb
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg & objDocument.Name & vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```

<br/>

The following example adds a new document to the **Documents** collection.

```vb
Dim objDocument As Document 
Set objDocument = Documents.Add 
With objDocument 
 .LayoutGuides.Columns = 4 
 .LayoutGuides.Rows = 9 
 .ActiveView.Zoom = pbZoomWholePage 
End With
```

<br/>

The following example displays the name of the first open publication.

```vb
If Documents.Count >= 1 Then 
 MsgBox Documents.Item(1).Name 
End If 

```

<br/>

The following example checks the name of each document in the **Documents** collection. If the name of a document is Sales.doc, an object variable `objSalesDoc` is set to that document in the **Documents** collection.

```vb
Dim objDocument As Document 
Dim objSalesDoc As Document 
For Each objDocument In Documents 
 If objDocument.Name = "Sales.doc" Then 
 Set objSalesDoc = objDocument 
 End If 
Next objDocument
```


## Methods

- [Add](Publisher.Documents.Add.md)

## Properties

- [Application](Publisher.Documents.Application.md)
- [Count](Publisher.Documents.Count.md)
- [Item](Publisher.Documents.Item.md)
- [Parent](Publisher.Documents.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]