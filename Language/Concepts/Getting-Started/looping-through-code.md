---
title: Looping through code (VBA)
keywords: vbcn6.chm1076674
f1_keywords:
- vbcn6.chm1076674
ms.prod: office
ms.assetid: 47d7ca2e-6103-c210-1c80-7ce76d854907
ms.date: 12/21/2018
localization_priority: Normal
---


# Looping through code

By using conditional statements and looping statements (also called control structures), you can write Visual Basic code that makes decisions and repeats actions. Another useful control structure, the **[With](../../reference/user-interface-help/with-statement.md)** statement, lets you run a series of statements without having to requalify an [object](../../Glossary/vbe-glossary.md#object).


## Use conditional statements to make decisions

Conditional statements evaluate whether a condition is **True** or **False**, and then specify one or more statements to run, depending on the result. Usually, a condition is an [expression](../../Glossary/vbe-glossary.md#expression) that uses a [comparison operator](../../Glossary/vbe-glossary.md#comparison-operator) to compare one value or [variable](../../Glossary/vbe-glossary.md#variable) with another.


### Choose a conditional statement to use

- [If...Then...Else](using-ifthenelse-statements.md): Branching when a condition is **True** or **False**   
- [Select Case](using-select-case-statements.md): Selecting a branch from a set of conditions
    

## Use loops to repeat code

Looping allows you to run a group of statements repeatedly. Some loops repeat statements until a condition is **False**; others repeat statements until a condition is **True**. There are also loops that repeat statements a specific number of times or for each object in a [collection](../../Glossary/vbe-glossary.md#collection).

### Choose a loop to use

- [Do...Loop](using-doloop-statements.md): Looping while or until a condition is **True**   
- [For...Next](using-fornext-statements.md): Using a counter to run statements a specified number of times   
- [For Each...Next](using-for-eachnext-statements.md): Repeating a group of statements for each object in a collection
    
## Run several statements on the same object

In Visual Basic, usually you must specify an object before you can run one of its [methods](../../Glossary/vbe-glossary.md#method) or change one of its [properties](../../Glossary/vbe-glossary.md#property). You can use the **With** statement to specify an object once for an entire series of statements.

- [With](using-with-statements.md): Running a series of statements on the same object
    
## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
