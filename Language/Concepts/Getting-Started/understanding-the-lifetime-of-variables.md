---
title: Understanding the lifetime of variables (VBA)
keywords: vbcn6.chm1076736
f1_keywords:
- vbcn6.chm1076736
ms.prod: office
ms.assetid: 018a61d5-4a0c-ac2e-6f06-50ba606034de
ms.date: 12/26/2018
localization_priority: Normal
---


# Understanding the lifetime of variables

The time during which a [variable](../../Glossary/vbe-glossary.md#variable) retains its value is known as its lifetime. The value of a variable may change over its lifetime, but it retains some value. When a variable loses [scope](../../Glossary/vbe-glossary.md#scope), it no longer has a value.

When a [procedure](../../Glossary/vbe-glossary.md#procedure) begins running, all variables are initialized. A numeric variable is initialized to zero, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with the character represented by the ASCII character code 0, or **Chr**( 0 ). [Variant](../../Glossary/vbe-glossary.md#variant-data-type) variables are initialized to [Empty](../../Glossary/vbe-glossary.md#empty). Each element of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) variable is initialized as if it were a separate variable.

When you declare an [object variable](../../Glossary/vbe-glossary.md#object-variable), space is reserved in memory, but its value is set to **[Nothing](../../reference/user-interface-help/nothing-keyword.md)** until you assign an object reference to it by using the **[Set](../../reference/user-interface-help/set-statement.md)** statement.

If the value of a variable isn't changed during the running of your code, it retains its initialized value until it loses scope.

A [procedure-level](../../Glossary/vbe-glossary.md#procedure-level) variable declared with the **[Dim](../../reference/user-interface-help/dim-statement.md)** statement retains a value until the procedure is finished running. If the procedure calls other procedures, the variable retains its value while those procedures are running as well.

If a procedure-level variable is declared with the **[Static](../../reference/user-interface-help/static-statement.md)** keyword, the variable retains its value as long as code is running in any [module](../../Glossary/vbe-glossary.md#module). When all code has finished running, the variable loses its scope and its value. Its lifetime is the same as a [module-level](../../Glossary/vbe-glossary.md#module-level) variable.

A module-level variable differs from a static variable. In a [standard module](../../Glossary/vbe-glossary.md#standard-module) or a [class module](../../Glossary/vbe-glossary.md#class-module), it retains its value until you stop running your code. In a class module, it retains its value as long as an instance of the class exists. Module-level variables consume memory resources until you reset their values, so use them only when necessary.

If you include the **Static** keyword before a **[Sub](../../reference/user-interface-help/sub-statement.md)** or **[Function](../../reference/user-interface-help/function-statement.md)** statement, the values of all the procedure-level variables in the procedure are preserved between calls.

## See also

- [Visual Basic reference](../../reference/user-interface-help/visual-basic-language-reference.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]