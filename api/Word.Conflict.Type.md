---
title: Conflict.Type property (Word)
keywords: vbawd10.chm78708740
f1_keywords:
- vbawd10.chm78708740
ms.prod: word
api_name:
- Word.Conflict.Type
ms.assetid: d2e5ad43-4b4b-8ce2-3aeb-453012759d9a
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflict.Type property (Word)

Returns the [WdRevisionType](Word.WdRevisionType.md)for the [Conflict](Word.Conflict.md) object. Read-only.


## Syntax

_expression_.**Type**

 _expression_ An expression that returns a '[Conflict](Word.Conflict.md)' object.


## Example

The following code example gets the [type](Word.Conflict.Type.md) of each conflict in the active document.


```vb
Dim con as Conflict 
 
For Each con in ActiveDocument.CoAuthoring.Conflicts 
 MsgBox con.Type 
Next con
```


## See also


[Conflict Object](Word.Conflict.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]