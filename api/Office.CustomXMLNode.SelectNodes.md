---
title: CustomXMLNode.SelectNodes method (Office)
keywords: vbaof11.chm294028
f1_keywords:
- vbaof11.chm294028
ms.prod: office
api_name:
- Office.CustomXMLNode.SelectNodes
ms.assetid: 443592af-a684-ee5e-98af-3e157f0f135e
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLNode.SelectNodes method (Office)

Selects a collection of nodes matching an XPath expression. This method differs from the **[CustomXMLPart.SelectNodes](office.customxmlpart.selectnodes.md)** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.


## Syntax

_expression_.**SelectNodes**(_XPath_)

_expression_ An expression that returns a **[CustomXMLNode](Office.CustomXMLNode.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains an XPath expression.|

## Return value

CustomXMLNodes


## Example

The following example demonstrates adding a custom XML part, selecting a part matching a namespace URI, and then selecting nodes within that part that match an XPath expression.


```vb
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Add a custom xml part. 
ActiveDocument.CustomXMLParts.Add "<supplier>" 
 
' Return the first custom xml part with the given namespace. 
Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")  
 
' Get all of the nodes matching an XPath expression. 
 Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")
```


## See also

- [CustomXMLNode object members](overview/library-reference/customxmlnode-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]