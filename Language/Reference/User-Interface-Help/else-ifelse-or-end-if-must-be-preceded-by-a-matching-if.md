---
title: Else If, Else, or End If must be preceded by a matching If
keywords: vblr6.chm1018974
f1_keywords:
- vblr6.chm1018974
ms.prod: office
ms.assetid: 7ec32fe8-a91f-e411-5c4f-ab1095b48d29
ms.date: 06/08/2017
localization_priority: Normal
---


# Else If, Else, or End If must be preceded by a matching If

**Else If**, **Else**, and **End If** are conditional compilation directives. This error has the following cause and solution:

An `#elseif`, `#else`, or `#endif` was detected that isn't preceded by a matching `#if` clause. Check to see if the intended `#if` has been separated from the clause in question by an intervening block or if the intended `#if` is preceded by a number sign (`#`) sign. If everything else is in order, place an `#if` clause in the appropriate position.
    
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]