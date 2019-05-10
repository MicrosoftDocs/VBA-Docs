---
title: ContentControl.Ungroup method (Word)
keywords: vbawd10.chm266534936
f1_keywords:
- vbawd10.chm266534936
ms.prod: word
api_name:
- Word.ContentControl.Ungroup
ms.assetid: 533e80a7-e2a0-ff46-3464-03e5de7faaf1
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControl.Ungroup method (Word)

Removes a group content control from a document so that its child content controls are no longer nested and can be freely edited.


## Syntax

_expression_.**Ungroup**

 _expression_ An expression that returns a [ContentControl](./Word.ContentControl.md) object.


## Remarks

This method fails if you run it on a control that is not of type  **wdContentControlGroup**.


## See also


[ContentControl Object](Word.ContentControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]