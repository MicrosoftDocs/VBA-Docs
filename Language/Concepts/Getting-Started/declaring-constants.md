---
title: Declaring constants (VBA)
keywords: vbcn6.chm1076698
f1_keywords:
- vbcn6.chm1076698
ms.prod: office
ms.assetid: c1b65bc4-1e94-828c-67bf-357a75261657
ms.date: 12/21/2018
localization_priority: Normal
---


# Declaring constants

By declaring a [constant](../../Glossary/vbe-glossary.md#constant), you can assign a meaningful name to a value. You use the **[Const](../../reference/user-interface-help/const-statement.md)** statement to declare a constant and set its value. After a constant is declared, it cannot be modified or assigned a new value.

You can declare a constant within a [procedure](../../Glossary/vbe-glossary.md#procedure) or at the top of a [module](../../Glossary/vbe-glossary.md#module), in the Declarations section. [Module-level](../../Glossary/vbe-glossary.md#module-level) constants are private by default. To declare a public module-level constant, precede the **Const** statement with the **Public** [keyword](../../Glossary/vbe-glossary.md#keyword). You can explicitly declare a private constant by preceding the **Const** statement with the **Private** keyword to make it easier to read and interpret your code. For more information, see [Understanding scope and visibility](understanding-scope-and-visibility.md).

The following example declares the **Public** constant `conAge` as an **Integer** and assigns it the value `34`.

```vb
Public Const conAge As Integer = 34
```

Constants can be declared as one of the following [data types](../../reference/user-interface-help/data-type-summary.md): **Boolean**, **Byte**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Date**, **String**, or **Variant**. Because you already know the value of a constant, you can specify the data type in a **Const** statement. 

You can declare several constants in one statement. To specify a data type, you must include the data type for each constant. 

In the following statement, the constants `conAge` and `conWage` are declared as **Integer**.

```vb
Const conAge As Integer = 34, conWage As Currency = 35000
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
