---
title: CustomXMLPrefixMappings.AddNamespace method (Office)
keywords: vbaof11.chm290004
f1_keywords:
- vbaof11.chm290004
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings.AddNamespace
ms.assetid: a4a58a81-3fdc-f808-ac19-0eb27e944f29
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPrefixMappings.AddNamespace method (Office)

Allows you to add a custom namespace/prefix mapping to use when querying an item.


## Syntax

_expression_.**AddNamespace**(_Prefix_, _NamespaceURI_)

_expression_ An expression that returns a **[CustomXMLPrefixMappings](Office.CustomXMLPrefixMappings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Prefix_|Required|**String**|Contains the prefix to add to the prefix mapping list.|
| _NamespaceURI_|Required|**String**|Contains the namespace to assign to the newly added prefix.|

## Remarks

If the prefix already exists in the **Namespace Manager**, this method will overwrite the meaning of that prefix except when the prefix is one added or used by the data store (**IXMLDataStore** interface) internally, in which case it will return an error.


## Example

The following example adds a prefix and namespace to a **CustomPrefixMappings** object.


```vb
Sub AddNamespacePrefix() 
  
    Dim objCustomPrefixMappings As  CustomPrefixMappings 
    Dim varCustomMapping As Variant 
 
    ' Adds a custom namespace. 
    varCustomMapping = objCustomPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")      
 
End Sub
```


## See also

- [CustomXMLPrefixMappings object members](overview/library-reference/customxmlprefixmappings-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]