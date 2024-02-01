---
title: CustomXMLSchemaCollection object (Office)
keywords: vbaof11.chm306000
f1_keywords:
- vbaof11.chm306000
api_name:
- Office.CustomXMLSchemaCollection
ms.assetid: 0ce1fe79-4287-303a-4205-586d8e116731
ms.date: 01/07/2019
ms.localizationpriority: medium
---


# CustomXMLSchemaCollection object (Office)

Represents a collection of **[CustomXMLSchema](Office.CustomXMLSchema.md)** objects attached to a data stream.


## Example

The following example adds a **CustomXMLSchema** object to a **CustomXMLSchemaCollection** object.


```vb
Dim SchemaCollection As CustomXMLSchemaCollection 
 
SchemaCollection.Add "https://tempuri.org/XMLSchema.xsd"
```


## See also

- [CustomXMLSchemaCollection object members](overview/library-reference/customxmlschemacollection-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]