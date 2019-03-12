---
title: Option Private statement (VBA)
keywords: vblr6.chm1011061
f1_keywords:
- vblr6.chm1011061
ms.prod: office
ms.assetid: bd4d8b8b-d513-62a0-7c78-45c15b462bdc
ms.date: 12/03/2018
localization_priority: Normal
---


# Option Private statement

When used in host applications that allow references across multiple [projects](../../Glossary/vbe-glossary.md#project), **Option Private Module** prevents a [module's](../../Glossary/vbe-glossary.md#module) contents from being referenced outside its project. In host applications that don't permit such references, for example, standalone versions of Visual Basic, **Option Private** has no effect.

## Syntax

**Option Private Module**

## Remarks

If used, the **Option Private** statement must appear at the [module level](../../Glossary/vbe-glossary.md#module-level), before any [procedures](../../Glossary/vbe-glossary.md#procedure).

When a module contains **Option Private Module**, the public parts, for example, [variables](../../Glossary/vbe-glossary.md#variable), [objects](../../Glossary/vbe-glossary.md#object), and [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type) declared at the module level, are still available within the [project](../../Glossary/vbe-glossary.md#project) containing the module, but they are not available to other applications or projects.

> [!NOTE] 
> **Option Private** is only useful for [host applications](../../Glossary/vbe-glossary.md#host-application) that support simultaneous loading of multiple projects and permit references between the loaded projects. For example, Microsoft Excel permits loading of multiple projects, and **Option Private Module** can be used to restrict cross-project visibility. Although Visual Basic permits loading of multiple projects, references between projects are never permitted in Visual Basic.


## Example

This example demonstrates the **Option Private** statement, which is used at module level to indicate that the entire module is private. With **Option Private Module**, module-level parts not declared **Private** are available to other modules in the project, but not to other projects or applications.


```vb
Option Private Module ' Indicates that module is private. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
