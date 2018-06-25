---
title: CustomXMLPrefixMapping Object (Office)
ms.prod: office
api_name:
- Office.CustomXMLPrefixMapping
ms.assetid: a657a760-cc52-5762-108e-2e95e9dba48f
ms.date: 06/08/2017
---


# CustomXMLPrefixMapping Object (Office)

Represents a namespace prefix.


## Example

The following example creates a  **CustomXMLPrefixMapping** object by adding a namespace and prefix to the **CustomXMLPrefixMapping** collection.


```vb
Dim objNamespace As CustomXMLPrefixMapping 
 
objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace") 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.CustomXMLPrefixMapping.Application.md)|
|[Creator](Office.CustomXMLPrefixMapping.Creator.md)|
|[NamespaceURI](Office.CustomXMLPrefixMapping.NamespaceURI.md)|
|[Parent](Office.CustomXMLPrefixMapping.Parent.md)|
|[Prefix](Office.CustomXMLPrefixMapping.Prefix.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
