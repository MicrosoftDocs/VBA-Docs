---
title: CustomXMLNode.SelectSingleNode method (Office)
keywords: vbaof11.chm294029
f1_keywords:
- vbaof11.chm294029
ms.prod: office
api_name:
- Office.CustomXMLNode.SelectSingleNode
ms.assetid: 630751f0-fe41-8f91-32d0-e266b3214cbf
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLNode.SelectSingleNode method (Office)

Selects a single node from a collection matching an XPath expression. This method differs from the **[CustomXMLPart.SelectSingleNode](office.customxmlpart.selectsinglenode.md)** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.


## Syntax

_expression_.**SelectSingleNode**(_XPath_)

_expression_ An expression that returns a **[CustomXMLNode](Office.CustomXMLNode.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains an XPath expression.|

## Return value

CustomXMLNode


## Remarks

The prefix mappings for the XPath expression are retrieved from the **[NamespaceManager](office.customxmlpart.namespacemanager.md)** property.


## Example

The following example demonstrates adding a custom XML part, selecting a part with a namespace URI, and then selecting a node within that part that matches an XPath expression.


```vb
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Add a custom xml part. 
ActiveDocument.CustomXMLParts.Add "<supplier>" 
 
' Return the first custom xml part with the given namespace. 
Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")         
 
' Get a node using XPath.                              
 Set cxn = cxp1(1).SelectSingleNode("//*[@supplierID = 1]") 

```


## See also

- [CustomXMLNode object members](overview/library-reference/customxmlnode-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]