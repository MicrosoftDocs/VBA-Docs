---
title: CustomXMLPart.AddNode method (Office)
keywords: vbaof11.chm295008
f1_keywords:
- vbaof11.chm295008
ms.prod: office
api_name:
- Office.CustomXMLPart.AddNode
ms.assetid: c316ebd0-e7e8-0ac2-603e-c298da23444d
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPart.AddNode method (Office)

Adds a node to the XML tree.


## Syntax

_expression_.**AddNode** (_Parent_, _Name_, _NamespaceURI_, _NextSibling_, _NodeType_, _NodeValue_)

_expression_ An expression that returns a **[CustomXMLPart](Office.CustomXMLPart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Parent_|Required|**CustomXMLNode**|Represents the node under which this node should be added. If adding an attribute, the parameter denotes the element that the attribute should be added to.|
| _Name_|Optional|**String**|Represents the base name of the node to be added.|
| _NamespaceURI_|Optional|**String**|Represents the namespace of the element to be appended. This parameter is required to append nodes of type **msoCustomXMLNodeElement** or **msoCustomXMLNodeAttribute**; otherwise, it is ignored.|
| _NextSibling_|Optional|**CustomXMLNode**|Represents the node which should become the next sibling of the new node. If not specified, the node is added to the end of the parent node's children. This parameter is ignored for additions of type **msoXMLNodeAttribute**. If the node is not a child of the parent, an error is displayed.|
| _NodeType_|Optional|**[MsoCustomXMLNodeType](office.msocustomxmlnodetype.md)**|Specifies the type of node to append. If the parameter is not specified, it is assumed to be of type **msoCustomXMLNodeElement**.|
| _NodeValue_|Optional|**String**|Used to set the value of the appended node for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.|

## Remarks

If the **AddNode** operation would result in an invalid tree structure, the append is not performed and an error message is displayed.


## Example

The following example demonstrates adding a node to a **CustomXMLPart** object.


```vb
Sub AddNodeCustomXmlParts() 
 
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
     
    With ActiveDocument 
        ' Add and populate a custom xml part 
        Set cxp1 = .CustomXMLParts.Add("<invoice />") 
         
        ' Set the parent node  
        Set cxn = cxp1.SelectSingleNode("/invoice") 
         
        ' Add a node under the parent node 
        cxp1.AddNode cxn, "upccode", "urn:invoice:namespace" 
 
    End With 
     
End Sub
```


## See also

- [CustomXMLPart object members](overview/library-reference/customxmlpart-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]