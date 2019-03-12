---
title: DocumentProperties object (Office)
keywords: vbaof11.chm250010
f1_keywords:
- vbaof11.chm250010
ms.prod: office
api_name:
- Office.DocumentProperties
ms.assetid: 90d42786-7d9a-b604-dbdf-88db41cbe69b
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperties object (Office)

A collection of **[DocumentProperty](office.documentproperty.md)** objects. Each **DocumentProperty** object represents a built-in or custom property of a container document.

## Remarks

Use the **Add** method to create a new custom property and add it to the **DocumentProperties** collection. You cannot use the **Add** method to create a built-in document property.

Use **BuiltinDocumentProperties(index)**, where _index_ is the index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. 

Use **CustomDocumentProperties(index)**, where _index_ is the number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property.


## See also

- [DocumentProperties object members](overview/library-reference/documentproperties-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
