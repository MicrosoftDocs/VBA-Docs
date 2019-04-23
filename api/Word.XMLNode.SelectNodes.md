---
title: XMLNode.SelectNodes method (Word)
keywords: vbawd10.chm37748755
f1_keywords:
- vbawd10.chm37748755
ms.prod: word
api_name:
- Word.XMLNode.SelectNodes
ms.assetid: a72d1693-a5da-bf97-179f-4fba2412c4ae
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.SelectNodes method (Word)

Returns an  **[XMLNodes](Word.XMLNodes.md)** collection that represents all the child elements that match the XPath parameter, in the order in which they appear within the specified XML element.


## Syntax

_expression_. `SelectNodes`( `_XPath_` , `_PrefixMapping_` , `_FastSearchSkippingTextNodes_` )

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|Specifies a valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.|
| _PrefixMapping_|Optional| **String**|Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.|
| _FastSearchSkippingTextNodes_|Optional| **Boolean**| **True** skips all text nodes while searching for the specified node. **False** includes text nodes in the search. Default value is **False**.|

## Return value

XMLNodes


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]