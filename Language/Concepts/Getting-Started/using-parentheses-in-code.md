---
title: Using parentheses in code (VBA)
keywords: vbcn6.chm1076685
f1_keywords:
- vbcn6.chm1076685
ms.prod: office
ms.assetid: 7894f174-ac01-dcc2-a30d-63d5c3625af6
ms.date: 12/26/2018
localization_priority: Normal
---


# Using parentheses in code

**[Sub](../../reference/user-interface-help/sub-statement.md)** procedures, built-in [statements](../../Glossary/vbe-glossary.md#statement), and some [methods](../../Glossary/vbe-glossary.md#method) don't return a value, so the [arguments](../../Glossary/vbe-glossary.md#argument) aren't enclosed in parentheses. For example:

```vb
MySub "stringArgument", integerArgument 

```


**[Function](../../reference/user-interface-help/function-statement.md)** procedures, built-in functions, and some methods do return a value, but you can ignore it. If you ignore the return value, don't include parentheses. Call the function just as you would call a **Sub** procedure. Omit the parentheses, list any arguments, and don't assign the function to a variable. For example:

```vb
MsgBox "Task Completed!", 0, "Task Box" 

```

To use the return value of a function, enclose the arguments in parentheses, as shown in the following example.

```vb
Answer3 = MsgBox("Are you happy with your salary?", 4, "Question 3") 

```

A statement in a **Sub** or **Function** procedure can pass values to a called procedure by using [named arguments](../../Glossary/vbe-glossary.md#named-argument). The guidelines for using parentheses apply, whether or not you use named arguments. When you use named arguments, you can list them in any order, and you can omit optional arguments. Named arguments are always followed by a colon and an equal sign (**:=**), and then the argument value.

The following example calls the **MsgBox** function by using named arguments, but it ignores the return value.

```vb
MsgBox Title:="Task Box", Prompt:="Task Completed!" 

```

The following example calls the **MsgBox** function by using named arguments and assigns the return value to the variable.

```vb
answer3 = MsgBox(Title:="Question 3", _ 
 Prompt:="Are you happy with your salary?", Buttons:=4) 

```


## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]