---
title: XMLNamespace.DefaultTransform property (Word)
keywords: vbawd10.chm2293766
f1_keywords:
- vbawd10.chm2293766
ms.prod: word
api_name:
- Word.XMLNamespace.DefaultTransform
ms.assetid: a43c9869-98f0-0a18-8e3c-eb4930553367
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNamespace.DefaultTransform property (Word)

 Returns an **[XSLTransform](Word.XSLTransform.md)** object that represents the default Extensible Stylesheet Language Transformation (XSLT) file to use when opening a document from an XML schema for a particular namespace.


## Syntax

_expression_. `DefaultTransform`

 _expression_ An expression that returns an '[XMLNamespace](Word.XMLNamespace.md)' object.


## Example

The following example returns the default XSLT for the first schema in the Schema Library that Microsoft Word will use to open XML files associated with that schema's namespace. This example assumes that the first schema has one or more applied XSLT files.


```vb
Dim objXSLT As XSLTransform 
 
Set objXSLT = Application.XMLNamespaces(1).DefaultTransform
```


## See also


[XMLNamespace Object](Word.XMLNamespace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]