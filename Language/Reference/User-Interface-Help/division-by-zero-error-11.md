---
title: Division by zero (Error 11)
keywords: vblr6.chm1011128
f1_keywords:
- vblr6.chm1011128
ms.prod: office
ms.assetid: 3c6783d9-24a4-ef25-fdab-9e26a08e35a9
ms.date: 06/08/2017
---


# Division by zero (Error 11)

Division by zero isn't possible. This error has the following cause and solution:



- The value of an [expression](../../Glossary/vbe-glossary.md#expression) being used as a divisor is zero.
    
    Check the spelling of [variables](../../Glossary/vbe-glossary.md#variable) in the expression. A misspelled variable name can implicitly create a numeric variable that is initialized to zero. Check previous operations on variables in the expression, especially those passed into the[procedure](../../Glossary/vbe-glossary.md#procedure) as[arguments](../../Glossary/vbe-glossary.md#argument) from other procedures.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

