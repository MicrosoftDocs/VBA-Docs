---
title: DocumentProperty object (Office)
keywords: vbaof11.chm250002
f1_keywords:
- vbaof11.chm250002
ms.prod: office
api_name:
- Office.DocumentProperty
ms.assetid: dd54ca3c-e0e2-4816-539a-17c5b4a928b1
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty object (Office)

Represents a custom or built-in document property of a container document. The **DocumentProperty** object is a member of the **[DocumentProperties](Office.DocumentProperties.md)** collection.


## Remarks

Use the Microsoft Word **Document.BuiltinDocumentProperties**(_index_) property, where _index_ is the name or index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. 

Use the Word **Document.CustomDocumentProperties**(_index_) property, where _index_ is the name or index number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property.

> [!NOTE] 
> Properties of type **msoPropertyTypeString** (**[MsoDocProperties](office.msodocproperties.md)**) are limited in length to 255 characters.


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]