---
title: Else clause must be preceded by a matching If
keywords: vblr6.chm1057022
f1_keywords:
- vblr6.chm1057022
ms.prod: office
ms.assetid: 4054aec1-ef5d-a939-3c7e-715d4dcde19f
ms.date: 06/08/2017
localization_priority: Normal
---


# Else clause must be preceded by a matching If

**Else** is a conditional compilation directive. This error has the following cause and solution:

An `#else` clause was detected that isn't preceded by a matching `#if` or `#elseif`. Check to see if a preceding `#if` has been separated from this `#else` by an `#endif`. Note that only one `#else` is permitted in each `#if` block, so two successive `#else` clauses cause this error.
    
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]