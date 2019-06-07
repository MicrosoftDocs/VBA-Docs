---
title: FindReplace.MatchCase property (Publisher)
keywords: vbapb10.chm8323080
f1_keywords:
- vbapb10.chm8323080
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchCase
ms.assetid: 4fabf2f8-f1e4-bc70-e8e6-96dd09cd23d8
ms.date: 06/07/2019
localization_priority: Normal
---


# FindReplace.MatchCase property (Publisher)

Sets or returns a **Boolean** that represents the case sensitivity of the search operation. Read/write.


## Syntax

_expression_.**MatchCase**

_expression_ A variable that represents a **[FindReplace](Publisher.FindReplace.md)** object.


## Return value

Boolean


## Remarks

The default value for **MatchCase** is **False**.


## Example

This example selects the first occurrence of the word "factory" regardless of case.

```vb
With ActiveDocument.Find 
 .Clear 
 .MatchCase = False 
 .FindText = "factory" 
 .Execute 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]