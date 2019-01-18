---
title: CustomXMLNode.RemoveChild method (Office)
keywords: vbaof11.chm294025
f1_keywords:
- vbaof11.chm294025
ms.prod: office
api_name:
- Office.CustomXMLNode.RemoveChild
ms.assetid: dc6c380a-6cfd-870a-9a31-d92aed1ae3e1
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLNode.RemoveChild method (Office)

Removes the specified child node from the tree.


## Syntax

_expression_.**RemoveChild**(_Child_)

_expression_ An expression that returns a **[CustomXMLNode](Office.CustomXMLNode.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Child_|Required|**CustomXMLNode**|Represents the child node of the context node.|

## Remarks

If the node specified in the _Child_ parameter is not a child of the context node, or if the action would result in an invalid tree, the removal is not performed and an error message is displayed.


## Example

The following example selects a custom part and then a node in that part. The code then removes a child of that node.


```vb
Dim cxp1 As CustomXMLPart 
 Dim cxn As CustomXMLNode 
 
 With ActiveDocument 
 
    ' Return the first part with the given root namespace. 
    Set cxp1 = .CustomXMLParts("urn:invoice:namespace")    
         
    ' Get node using XPath expression.                              
    Set cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")  
 
    ' Remove a child node. 
    cxn.RemoveChild(cxn.SelectSingleNode("//discount"))   
        
End With     

```


## See also

- [CustomXMLNode object members](overview/library-reference/customxmlnode-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]