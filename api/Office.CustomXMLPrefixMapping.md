---
title: CustomXMLPrefixMapping object (Office)
ms.prod: office
api_name:
- Office.CustomXMLPrefixMapping
ms.assetid: a657a760-cc52-5762-108e-2e95e9dba48f
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPrefixMapping object (Office)

Represents a namespace prefix.


## Example

The following example creates a **CustomXMLPrefixMapping** object by adding a namespace and prefix to the **CustomXMLPrefixMapping** collection.


```vb
Dim objNamespace As CustomXMLPrefixMapping 
 
objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace") 

```


## See also

- [CustomXMLPrefixMapping object members](overview/library-reference/customxmlprefixmapping-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]