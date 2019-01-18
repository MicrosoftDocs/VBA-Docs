---
title: Row.NestingLevel property (Word)
keywords: vbawd10.chm156237930
f1_keywords:
- vbawd10.chm156237930
ms.prod: word
api_name:
- Word.Row.NestingLevel
ms.assetid: ad67f444-7d9c-a749-0cff-811aa5f30697
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.NestingLevel property (Word)

Returns the nesting level of the specified table row. Read-only  **Long**.


## Syntax

 _expression_. `NestingLevel`

 _expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Remarks

The outermost table has a nesting level of 1. The nesting level of each successively nested table is one higher than the previous table.


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]