---
title: Tables.NestingLevel property (Word)
keywords: vbawd10.chm156041316
f1_keywords:
- vbawd10.chm156041316
ms.prod: word
api_name:
- Word.Tables.NestingLevel
ms.assetid: 50a0860d-9ad2-8fe3-4cc7-108527d72084
ms.date: 06/08/2017
localization_priority: Normal
---


# Tables.NestingLevel property (Word)

Returns the nesting level of the specified tables. Read-only  **Long**.


## Syntax

 _expression_. `NestingLevel`

 _expression_ Required. A variable that represents a '[Tables](Word.tables.md)' collection.


## Remarks

The outermost table has a nesting level of 1. The nesting level of each successively nested table is one higher than the previous table.


## See also


[Tables Collection Object](Word.tables.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]