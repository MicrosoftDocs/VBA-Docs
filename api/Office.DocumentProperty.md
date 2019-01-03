---
title: DocumentProperty object (Office)
keywords: vbaof11.chm250002
f1_keywords:
- vbaof11.chm250002
ms.prod: office
api_name:
- Office.DocumentProperty
ms.assetid: dd54ca3c-e0e2-4816-539a-17c5b4a928b1
ms.date: 06/08/2017
---


# DocumentProperty object (Office)

Represents a custom or built-in document property of a container document. The **DocumentProperty** object is a member of the **DocumentProperties** collection.


## Remarks

Use the Microsoft Word **Document.BuiltinDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. Use the Microsoft Word **Document.CustomDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property. The following list contains the names of all the available built-in document properties:


> [!NOTE] 
> Properties of type **msoPropertyTypeString** are limited in length to 255 characters.


## See also

- [Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
- [DocumentProperty Object Members](./overview/Library-Reference/documentproperty-members-office.md)

