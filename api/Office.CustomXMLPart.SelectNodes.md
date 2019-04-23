---
title: CustomXMLPart.SelectNodes method (Office)
keywords: vbaof11.chm295012
f1_keywords:
- vbaof11.chm295012
ms.prod: office
api_name:
- Office.CustomXMLPart.SelectNodes
ms.assetid: c220c535-ac3f-cdba-5b1b-b608ed2eb8e4
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPart.SelectNodes method (Office)

Selects a collection of nodes from a custom XML part.


## Syntax

_expression_.**SelectNodes**(_XPath_)

_expression_ An expression that returns a **[CustomXMLPart](Office.CustomXMLPart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains the XPath expression.|

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

- [CustomXMLPart object members](overview/library-reference/customxmlpart-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]