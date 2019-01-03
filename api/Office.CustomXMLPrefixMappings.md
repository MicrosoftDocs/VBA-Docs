---
title: CustomXMLPrefixMappings object (Office)
keywords: vbaof11.chm290000
f1_keywords:
- vbaof11.chm290000
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings
ms.assetid: 7da5e1df-a436-ab54-4ea0-270f3edaf240
ms.date: 06/08/2017
---


# CustomXMLPrefixMappings object (Office)

Represents a collection of  **CustomXMLPrefixMapping** objects.


## Example

The following example creates a  **CustomXMLPrefixMapping** object by adding a namespace and prefix to the **CustomXMLPrefixMapping** collection.


```vb
Dim objNamespace As CustomXMLPrefixMapping 
 
objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")
```


## Methods



|**Name**|
|:-----|
|[AddNamespace](Office.CustomXMLPrefixMappings.AddNamespace.md)|
|[LookupNamespace](Office.CustomXMLPrefixMappings.LookupNamespace.md)|
|[LookupPrefix](Office.CustomXMLPrefixMappings.LookupPrefix.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.CustomXMLPrefixMappings.Application.md)|
|[Count](Office.CustomXMLPrefixMappings.Count.md)|
|[Creator](Office.CustomXMLPrefixMappings.Creator.md)|
|[Item](Office.CustomXMLPrefixMappings.Item.md)|
|[Parent](Office.CustomXMLPrefixMappings.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
