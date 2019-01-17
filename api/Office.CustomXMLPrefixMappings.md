---
title: CustomXMLPrefixMappings object (Office)
keywords: vbaof11.chm290000
f1_keywords:
- vbaof11.chm290000
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings
ms.assetid: 7da5e1df-a436-ab54-4ea0-270f3edaf240
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPrefixMappings object (Office)

Represents a collection of **[CustomXMLPrefixMapping](Office.CustomXMLPrefixMapping.md)** objects.


## Example

The following example creates a **CustomXMLPrefixMapping** object by adding a namespace and prefix to the **CustomXMLPrefixMapping** collection.


```vb
Dim objNamespace As CustomXMLPrefixMapping 
 
objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")
```


## See also

- [CustomXMLPrefixMappings object members](overview/library-reference/customxmlprefixmappings-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]