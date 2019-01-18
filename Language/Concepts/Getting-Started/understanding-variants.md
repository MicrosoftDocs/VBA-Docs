---
title: Understanding Variants (VBA)
keywords: vbcn6.chm1076678
f1_keywords:
- vbcn6.chm1076678
ms.prod: office
ms.assetid: 0f8d3917-0ca3-0a67-2c3d-48883f4a24f1
ms.date: 12/26/2018
localization_priority: Normal
---


# Understanding Variants

The **[Variant](../../reference/user-interface-help/variant-data-type.md)** data type is automatically specified if you don't specify a [data type](../../Glossary/vbe-glossary.md#data-type) when you declare a [constant](../../Glossary/vbe-glossary.md#constant), [variable](../../Glossary/vbe-glossary.md#variable), or [argument](../../Glossary/vbe-glossary.md#argument). 

Variables declared as the **Variant** data type can contain string, date, time, Boolean, or numeric values, and can convert the values that they contain automatically. Numeric **Variant** values require 16 bytes of memory (which is significant only in large [procedures](../../Glossary/vbe-glossary.md#procedure) or complex [modules](../../Glossary/vbe-glossary.md#module)), and they are slower to access than explicitly typed variables of any other type. You rarely use the **Variant** data type for a constant. String **Variant** values require 22 bytes of memory.

The following statements create **Variant** variables:

```vb
Dim myVar 
Dim yourVar As Variant 
theVar = "This is some text." 

```

The last statement does not explicitly declare the variable, but rather declares the variable implicitly, or automatically. Variables that are declared implicitly are specified as the **Variant** data type.

> [!TIP] 
> If you specify a data type for a variable or argument, and then use the wrong data type, a data type error will occur. To avoid data type errors, either use only implicit variables (the **Variant** data type) or explicitly declare all your variables and specify a data type. The latter method is preferred.


## See also

- [Visual Basic reference](../../reference/user-interface-help/visual-basic-language-reference.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]