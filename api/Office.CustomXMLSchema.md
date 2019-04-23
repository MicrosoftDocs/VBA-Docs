---
title: CustomXMLSchema object (Office)
keywords: vbaof11.chm291000
f1_keywords:
- vbaof11.chm291000
ms.prod: office
api_name:
- Office.CustomXMLSchema
ms.assetid: 9110da6c-fc54-98b2-7e5e-e6d4c21712ad
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLSchema object (Office)

Represents a schema in a **CustomXMLSchemaCollection** collection.


## Example

The following example adds a **CustomXMLSchema** object to a **CustomXMLSchemaCollection** object.


```vb
Dim SchemaCollection As CustomXMLSchemaCollection 
 
SchemaCollection.Add "https://tempuri.org/XMLSchema.xsd" 

```


## See also

- [CustomXMLSchema object members](overview/library-reference/customxmlschema-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]