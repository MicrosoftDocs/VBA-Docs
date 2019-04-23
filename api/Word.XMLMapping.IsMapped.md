---
title: XMLMapping.IsMapped property (Word)
keywords: vbawd10.chm199688192
f1_keywords:
- vbawd10.chm199688192
ms.prod: word
api_name:
- Word.XMLMapping.IsMapped
ms.assetid: e78ae752-1f8f-5f18-0755-97ec10ab68ec
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLMapping.IsMapped property (Word)

Returns a  **Boolean** that represents whether the content control in the document is mapped to an XML node in the document's XML data store. Read-only.


## Syntax

_expression_. `IsMapped`

 _expression_ An expression that returns an '[XMLMapping](Word.XMLMapping.md)' object.


## Example

The following example deletes the XML mapping for all content controls in the active document that are currently mapped.


```vb
Dim objCC As ContentControl 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.XMLMapping.IsMapped Then 
 objCC.XMLMapping.Delete 
 End If 
Next
```


## See also


[XMLMapping Object](Word.XMLMapping.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]