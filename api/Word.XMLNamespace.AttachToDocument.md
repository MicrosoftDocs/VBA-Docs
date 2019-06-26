---
title: XMLNamespace.AttachToDocument method (Word)
keywords: vbawd10.chm2293860
f1_keywords:
- vbawd10.chm2293860
ms.prod: word
api_name:
- Word.XMLNamespace.AttachToDocument
ms.assetid: 18af2ed2-2806-401a-4cca-9d8746f25082
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNamespace.AttachToDocument method (Word)

Attaches an XML schema to a document.


## Syntax

_expression_. `AttachToDocument`( `_Document_` )

 _expression_ An expression that represents a '[XMLNamespace](Word.XMLNamespace.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Document_|Required| **Document**|The document to which to attach the specified XML schema.|

## Example

The following example adds the SimpleSample schema to the Schema Library and then attaches it to the active document.


> [!NOTE] 
> The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


```vb
Dim objSchema As XMLNamespace 
 
Set objSchema = Application.XMLNamespaces _ 
 .Add("c:\schemas\simplesample.xsd") 
 
objSchema.AttachToDocument ActiveDocument
```


## See also


[XMLNamespace Object](Word.XMLNamespace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]