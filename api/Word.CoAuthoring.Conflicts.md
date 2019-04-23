---
title: CoAuthoring.Conflicts property (Word)
keywords: vbawd10.chm254869511
f1_keywords:
- vbawd10.chm254869511
ms.prod: word
api_name:
- Word.CoAuthoring.Conflicts
ms.assetid: bd6aab5d-5342-ee1b-c5af-1f67753d55fc
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthoring.Conflicts property (Word)

Returns a  **[Conflicts](Word.Conflicts.md)** collection that represents all the conflicts in a document. Read-only.


## Syntax

_expression_. `Conflicts`

 _expression_ An expression that returns a '[CoAuthoring](Word.CoAuthoring.md)' object.


## Example

The following code example gets the type of each conflict in the active document. The  **[Type](Word.Conflict.Type.md)** property uses the **[WdRevisionType](Word.WdRevisionType.md)** enumeration to specify the conflict type.


```vb
Dim conf As Conflict 
 
For Each conf In ActiveDocument.CoAuthoring.Conflicts 
    MsgBox conf.Type 
Next conf 

```


## See also


[CoAuthoring Object](Word.CoAuthoring.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]