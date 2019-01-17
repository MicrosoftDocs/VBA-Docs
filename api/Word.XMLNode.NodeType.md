---
title: XMLNode.NodeType property (Word)
keywords: vbawd10.chm37748748
f1_keywords:
- vbawd10.chm37748748
ms.prod: word
api_name:
- Word.XMLNode.NodeType
ms.assetid: 0df07d30-e7ae-44e6-3372-ccece783a3fc
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.NodeType property (Word)

Returns a  **[WdXMLNodeType](./overview/Word.md)** constant that represents the type of node.


## Syntax

 _expression_. `NodeType`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Remarks

An  **XMLNode** object can be either an XML element or an attribute of an element. Use the **NodeType** property to determine which type of node you are working with, so that you do not attempt to perform invalid operations on the node. For example, the **[Attributes](Word.XMLNode.Attributes.md)** property applies only to element nodes, although it appears in the list of available properties for the **XMLNode** object.


## Example

The following example adds the author attribute to the book element in the active document and then sets the value of the attribute.


```vb
Sub AddIDAttribute() 
 Dim objElement As XMLNode 
 Dim objAttribute As XMLNode 
 
 For Each objElement In ActiveDocument.XMLNodes 
 If objElement.NodeType = wdXMLNodeElement Then 
 If objElement.BaseName = "book" Then 
 
 Set objAttribute = objElement.Attributes _ 
 .Add("author", objElement.NamespaceURI) 
 
 objAttribute.NodeValue = "David Barber" 
 
 Exit For 
 End If 
 End If 
 Next 
End Sub
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]