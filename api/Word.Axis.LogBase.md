---
title: Axis.LogBase property (Word)
keywords: vbawd10.chm113049622
f1_keywords:
- vbawd10.chm113049622
ms.prod: word
api_name:
- Word.Axis.LogBase
ms.assetid: bf6be786-60e4-789f-792b-f866d88d7066
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.LogBase property (Word)

Returns or sets the base of the logarithm when you are using log scales. Read/write  **Double**.


## Syntax

_expression_.**LogBase**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Attempting to set this property to a value less than or equal to zero (0) raises an error. The default value is 10.


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]