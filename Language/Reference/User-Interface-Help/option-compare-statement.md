---
title: Option Compare statement (VBA)
keywords: vblr6.chm1008991
f1_keywords:
- vblr6.chm1008991
ms.prod: office
ms.assetid: 9332562c-451e-50df-198a-21902fadac9c
ms.date: 12/03/2018
localization_priority: Normal
---


# Option Compare statement

Used at the [module level](../../Glossary/vbe-glossary.md#module-level) to declare the default comparison method to use when string data is compared.

## Syntax

**Option Compare** { **Binary** | **Text** | **Database** }

## Remarks

If used, the **Option Compare** statement must appear in a [module](../../Glossary/vbe-glossary.md#module) before any [procedures](../../Glossary/vbe-glossary.md#procedure).

The **Option Compare** statement specifies the [string comparison](../../Glossary/vbe-glossary.md#string-comparison) method (**Binary**, **Text**, or **Database**) for a module. If a module doesn't include an **Option Compare** statement, the default text comparison method is **Binary**.

- **Option Compare Binary** results in string comparisons based on a [sort order](../../Glossary/vbe-glossary.md#sort-order) derived from the internal binary representations of the characters. In Microsoft Windows, sort order is determined by the code page. A typical binary sort order is shown in the following example:

  ```vb
   A < B < E < Z < a < b < e < z < À < Ê < Ø < à < ê < ø 
  
  ```

- **Option Compare Text** results in string comparisons based on a case-insensitive text sort order determined by your system's [locale](../../Glossary/vbe-glossary.md#locale). When the same characters are sorted by using **Option Compare Text**, the following text sort order is produced:

  ```vb
    (A=a) < ( À=à) < (B=b) < (E=e) < (Ê=ê) < (Z=z) < (Ø=ø) 

  ```

- **Option Compare Database** can only be used within Microsoft Access. This results in string comparisons based on the sort order determined by the locale ID of the database where the string comparisons occur.

## Example

This example uses the **Option Compare** statement to set the default string comparison method. The **Option Compare** statement is used at the module level only.


```vb
' Set the string comparison method to Binary. 
Option Compare Binary ' That is, "AAA" is less than "aaa". 
' Set the string comparison method to Text. 
Option Compare Text ' That is, "AAA" is equal to "aaa". 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
