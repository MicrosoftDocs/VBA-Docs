---
title: Document.CustomDocumentProperties Property (Word)
keywords: vbawd10.chm158007298
f1_keywords:
- vbawd10.chm158007298
ms.prod: word
api_name:
- Word.Document.CustomDocumentProperties
ms.assetid: 4f8ac449-b9b3-45a0-7962-df7252067e67
ms.date: 06/08/2017
---


# Document.CustomDocumentProperties Property (Word)

Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified document.


## Syntax

 _expression_. `CustomDocumentProperties`

 _expression_ Required. A variable that represents a '[Document](Word.Document.md)' object.


## Remarks

Use the  **BuiltInDocumentProperties** property to return the collection of built-in document properties.

Properties of type  **msoPropertyTypeString** cannot exceed 255 characters in length.

For information about returning a single member of a collection, see [Returning an Object from a Collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example inserts a list of custom built-in properties at the end of the active document.


```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
For Each prop In ActiveDocument.CustomDocumentProperties 
 With myRange 
 .InsertParagraphAfter 
<<<<<<< HEAD
 .InsertAfter prop.Name &; "= " 
=======
 .InsertAfter prop.Name & "= " 
>>>>>>> master
 .InsertAfter prop.Value 
 End With 
Next
```

This example adds a custom built-in property to Sales.doc.




```vb
thename = InputBox("Please type your name", "Name") 
Documents("Sales.doc").CustomDocumentProperties.Add _ 
 Name:="YourName", LinkToContent:=False, Value:=thename, _ 
 Type:=msoPropertyTypeString
```


## See also


[Document Object](Word.Document.md)

