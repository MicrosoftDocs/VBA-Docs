---
title: + operator
keywords: vblr6.chm1008845
f1_keywords:
- vblr6.chm1008845
ms.prod: office
ms.assetid: fa6ce797-a49c-af99-4ab5-112056c2a584
ms.date: 11/19/2018
localization_priority: Normal
---


# + operator

Used to sum two numbers.

## Syntax

_result_ = _expression1 + expression2_

The **+** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _expression1_|Required; any [expression](../../Glossary/vbe-glossary.md#expression).|
| _expression2_|Required; any expression.|

## Remarks

When you use the **+** operator, you may not be able to determine whether addition or string concatenation will occur. Use the **[&](ampersand-operator.md)** operator for concatenation to eliminate ambiguity and provide self-documenting code.

If at least one expression is not a [Variant](../../Glossary/vbe-glossary.md#variant-data-type), the following rules apply.

|If|Then|
|:-----|:-----|
|Both expressions are [numeric data types](../../Glossary/vbe-glossary.md#numeric-data-type) ([Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Date](../../Glossary/vbe-glossary.md#date-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), or [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type))|Add.|
|Both expressions are [String](../../Glossary/vbe-glossary.md#string-data-type)|Concatenate.|
|One expression is a numeric data type and the other is any **Variant** except [Null](../../Glossary/vbe-glossary.md#null)|Add.|
|One expression is a **String** and the other is any **Variant** except **Null**|Concatenate.|
|One expression is an [Empty](../../Glossary/vbe-glossary.md#empty) **Variant**|Return the remaining expression unchanged as _result_.|
|One expression is a numeric data type and the other is a **String**|A `Type mismatch`error occurs.|
|Either expression is **Null**| _result_ is **Null**.|

<br/>

If both expressions are **Variant** expressions, the following rules apply:

|If|Then|
|:-----|:-----|
|Both **Variant** expressions are numeric|Add.|
|Both **Variant** expressions are strings|Concatenate.|
|One **Variant** expression is numeric and the other is a string|Add.|

<br/>

For simple arithmetic addition involving only expressions of numeric data types, the [data type](../../Glossary/vbe-glossary.md#data-type) of _result_ is usually the same as that of the most precise expression. The order of precision, from least to most precise, is **Byte**, **Integer**, **Long**, **Single**, **Double**, **Currency**, and **Decimal**. The following are exceptions to this order.

|If|Then _result_ is|
|:-----|:-----|
|A **Single** and a **Long** are added |A **Double**.|
|The data type of _result_ is a **Long**, **Single**, or **Date** variant that overflows its legal range|Converted to a **Double** variant.|
|The data type of _result_ is a **Byte** variant that overflows its legal range|Converted to an **Integer** variant.|
|The data type of _result_ is an **Integer** variant that overflows its legal range|Converted to a **Long** variant.|
|A **Date** is added to any data type|A **Date**.|


<br/>

If one or both expressions are **Null** expressions, _result_ is **Null**. If both expressions are **Empty**, _result_ is an **Integer**. However, if only one expression is **Empty**, the other expression is returned unchanged as _result_.

> [!NOTE] 
> The order of precision used by addition and subtraction is not the same as the order of precision used by multiplication.

## Example

This example uses the **+** operator to sum numbers. The **+** operator can also be used to concatenate strings. However, to eliminate ambiguity, you should use the **&** operator instead. If the components of an expression created with the **+** operator include both strings and numerics, the arithmetic result is assigned. If the components are exclusively strings, the strings are concatenated.


```vb
Dim MyNumber, Var1, Var2
MyNumber = 2 + 2    ' Returns 4.
MyNumber = 4257.04 + 98112    ' Returns 102369.04.

Var1 = "34": Var2 = 6    ' Initialize mixed variables.
MyNumber = Var1 + Var2    ' Returns 40.

Var1 = "34": Var2 = "6"    ' Initialize variables with strings.
MyNumber = Var1 + Var2    ' Returns "346" (string concatenation).

```


## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]