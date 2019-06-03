---
title: CustomXMLPart.SelectSingleNode method (Office)
keywords: vbaof11.chm295013
f1_keywords:
- vbaof11.chm295013
ms.prod: office
api_name:
- Office.CustomXMLPart.SelectSingleNode
ms.assetid: 2bd4c25b-d4e6-08db-b2ce-c74adf16336f
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPart.SelectSingleNode method (Office)

Selects a single node within a custom XML part matching an XPath expression.


## Syntax

_expression_.**SelectSingleNode**(_XPath_)

_expression_ An expression that returns a **[CustomXMLPart](Office.CustomXMLPart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains an XPath expression.|

## Return value

CustomXMLNode


## Example

The following example demonstrates adding a custom XML part, selecting a part with a namespace URI, and then selecting a node within that part that matches an XPath expression. 


```vb

Dim cxp1 As CustomXMLPart
Dim cxn As CustomXMLNode

' Add a custom XML part.
ActiveDocument.CustomXMLParts.Add ( _
    "<suppliers>" & _
    "<supplier ID='1'>Contoso</supplier>" & _
    "<supplier ID='2'>Wingtip Toys</supplier>" & _
    "</suppliers>")

' Return the last custom XML part added to the document.
Set cxp1 = ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count)

' Get a node using XPath.
Set cxn = cxp1.SelectSingleNode("//supplier[@ID=1]")

' Display the node value 'Contoso'.
MsgBox cxn.NodeValue


```


## See also

- [CustomXMLPart object members](overview/library-reference/customxmlpart-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]