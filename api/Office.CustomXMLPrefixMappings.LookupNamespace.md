---
title: CustomXMLPrefixMappings.LookupNamespace method (Office)
keywords: vbaof11.chm290005
f1_keywords:
- vbaof11.chm290005
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings.LookupNamespace
ms.assetid: 33a8f054-0e67-0c9e-ce4b-c9d3360df1a6
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPrefixMappings.LookupNamespace method (Office)

Allows you to get the namespace corresponding to the specified prefix.


## Syntax

_expression_.**LookupNamespace**(_Prefix_)

_expression_ An expression that returns a **[CustomXMLPrefixMappings](Office.CustomXMLPrefixMappings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Prefix_|Required|**String**|Contains a prefix in the prefix mapping list.|

## Return value

String


## Remarks

If no namespace is assigned to the requested prefix, the method returns an empty string ("").


## Example

The following example retrieves the namespace corresponding to the prefix argument.


```vb
Dim objCustomPrefixMappings As  CustomPrefixMappings 
    Dim strNamespace As String 
 
    ' Gets the namespace corresponding to the prefix. 
   strNamespace = objCustomPrefixMappings.LookupNamespace("xs")
```


## See also

- [CustomXMLPrefixMappings object members](overview/library-reference/customxmlprefixmappings-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]