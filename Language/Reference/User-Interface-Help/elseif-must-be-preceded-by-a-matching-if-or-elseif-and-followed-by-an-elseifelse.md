---
title: ElseIf must be preceded by a matching If or ElseIf and followed by an ElseIf, Else, or EndIf
keywords: vblr6.chm1057023
f1_keywords:
- vblr6.chm1057023
ms.prod: office
ms.assetid: f6ded2b0-e05b-643e-6599-9cf3ed592a7d
ms.date: 06/08/2017
localization_priority: Normal
---


# ElseIf must be preceded by a matching If or ElseIf and followed by an ElseIf, Else, or EndIf

**ElseIf** is a conditional compilation directive. This error has the following causes and solutions:

- An `#elseif` has been detected that isn't preceded by an `#if` or `#elseif`. Place an `#if` statement before the `#elseif` or remove an incorrectly placed preceding `#endif`.
    
- An `#elseif` has been detected that is preceded by an `#else` or `#endif`. Appropriately terminate the preceding `#if` block, or change the preceding `#else` to an `#elseif`.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]