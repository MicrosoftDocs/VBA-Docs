---
title: XMLMapping.CustomXMLPart property (Word)
keywords: vbawd10.chm199688193
f1_keywords:
- vbawd10.chm199688193
ms.prod: word
api_name:
- Word.XMLMapping.CustomXMLPart
ms.assetid: a9eac7d6-0088-7251-e0b2-fef529fee278
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLMapping.CustomXMLPart property (Word)

Returns a  **CustomXMLPart** object that represents the custom XML part to which the content control in the document maps.


## Syntax

_expression_. `CustomXMLPart`

 _expression_ An expression that returns an '[XMLMapping](Word.XMLMapping.md)' object.


## Example

The following example accesses the first content control in the active document and the custom XML part to which it is mapped, and then sets the text value of one of the XML nodes contained within the custom XML part.


> [!NOTE] 
> This example assumes that at least one content control in the active document is mapped to a custom XML part that contains the XML nodes specified in the XPath string.


```vb
Dim objCC As ContentControl 
Dim objPart As CustomXMLPart 
Dim objNode As CustomXMLNode 
 
Set objCC = ActiveDocument.ContentControls(1) 
Set objPart = objCC.XMLMapping.CustomXMLPart 
Set objNode = objPart.SelectSingleNode("/books/book/title") 
objNode.Text = "Mystery of the Empty Chair"
```


## See also


[XMLMapping Object](Word.XMLMapping.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]