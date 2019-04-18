---
title: Comparison operators
keywords: vblr6.chm1008875
f1_keywords:
- vblr6.chm1008875
ms.prod: office
ms.assetid: 9c254e88-5641-ea7d-b99a-cb614c3095a7
ms.date: 11/19/2018
localization_priority: Normal
---


# Comparison operators 

Used to compare [expressions](../../Glossary/vbe-glossary.md#expression).

## Syntax

_result_ = _expression1_ _comparisonoperator_ _expression2_ <br/>
_result_ = _object1_ **Is** _object2_ <br/>
_result_ = _string_ **Like** _pattern_ 

[Comparison operators](../../Glossary/vbe-glossary.md#comparison-operator) have these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _expression_|Required; any expression.|
| _comparisonoperator_|Required; any comparison operator.|
| _object_|Required; any object name.|
| _string_|Required; any [string expression](../../Glossary/vbe-glossary.md#string-expression).|
| _pattern_|Required; any string expression or range of characters.|

## Remarks

The following table contains a list of the comparison operators and the conditions that determine whether _result_ is **True**, **False**, or [Null](../../Glossary/vbe-glossary.md#null).


|Operator|True if|False if|Null if|
|:-----|:-----|:-----|:-----|
|`<` (Less than)| _expression1_ < _expression2_| _expression1_ >= _expression2_| _expression1_ or _expression2_ = **Null**|
|`<=` (Less than or equal to)| _expression1_ <= _expression2_| _expression1_ > _expression2_| _expression1_ or _expression2_ = **Null**|
|`>` (Greater than)| _expression1_ > _expression2_| _expression1_ <= _expression2_| _expression1_ or _expression2_ = **Null**|
|`>=` (Greater than or equal to)| _expression1_ >= _expression2_| _expression1_ < _expression2_| _expression1_ or _expression2_ = **Null**|
|`=` (Equal to)| _expression1_ = _expression2_| _expression1_ <> _expression2_| _expression1_ or _expression2_ = **Null**|
|`<>` (Not equal to)| _expression1_ <> _expression2_| _expression1_ = _expression2_| _expression1_ or _expression2_ = **Null**|

> [!NOTE] 
> The **Is** and **Like** operators have specific comparison functionality that differs from the operators in the table.

When comparing two expressions, you may not be able to easily determine whether the expressions are being compared as numbers or as strings. The following table shows how the expressions are compared or the result when either expression is not a [Variant](../../Glossary/vbe-glossary.md#variant-data-type).

|If|Then|
|:-----|:-----|
|Both expressions are [numeric data types](../../Glossary/vbe-glossary.md#numeric-data-type) ([Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Date](../../Glossary/vbe-glossary.md#date-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), or [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type))|Perform a numeric comparison.|
|Both expressions are [String](../../Glossary/vbe-glossary.md#string-data-type)|Perform a [string comparison](../../Glossary/vbe-glossary.md#string-comparison).|
|One expression is a numeric data type and the other is a **Variant** that is, or can be, a number|Perform a numeric comparison.|
|One expression is a numeric data type and the other is a string **Variant** that can't be converted to a number|A  `Type Mismatch` error occurs.|
|One expression is a **String** and the other is any **Variant** except a **Null**|Perform a string comparison.|
|One expression is [Empty](../../Glossary/vbe-glossary.md#empty) and the other is a numeric data type|Perform a numeric comparison, using 0 as the **Empty** expression.|
|One expression is **Empty** and the other is a **String**|Perform a string comparison, using a zero-length string ("") as the **Empty** expression.|

<br/>

If _expression1_ and _expression2_ are both **Variant** expressions, their underlying type determines how they are compared. The following table shows how the expressions are compared or the result from the comparison, depending on the underlying type of the **Variant**.

|If|Then|
|:-----|:-----|
|Both **Variant** expressions are numeric|Perform a numeric comparison.|
|Both **Variant** expressions are strings|Perform a string comparison.|
|One **Variant** expression is numeric and the other is a string|The numeric expression is less than the string expression.|
|One **Variant** expression is **Empty** and the other is numeric|Perform a numeric comparison, using 0 as the **Empty** expression.|
|One **Variant** expression is **Empty** and the other is a string|Perform a string comparison, using a zero-length string ("") as the **Empty** expression.|
|Both **Variant** expressions are **Empty**|The expressions are equal.|

When a **Single** is compared to a **Double**, the **Double** is rounded to the precision of the **Single**.
If a **Currency** is compared with a **Single** or **Double**, the **Single** or **Double** is converted to a **Currency**. 

Similarly, when a **Decimal** is compared with a **Single** or **Double**, the **Single** or **Double** is converted to a **Decimal**. For **Currency**, any fractional value less than .0001 may be lost; for **Decimal**, any fractional value less than 1E-28 may be lost, or an overflow error can occur. Such fractional value loss may cause two values to compare as equal when they are not.

## Example

This example shows various uses of comparison operators, which you use to compare expressions.

```vb
Dim MyResult, Var1, Var2
MyResult = (45 < 35)    ' Returns False.
MyResult = (45 = 45)    ' Returns True.
MyResult = (4 <> 3)    ' Returns True.
MyResult = ("5" > "4")    ' Returns True.

Var1 = "5": Var2 = 4    ' Initialize variables.
MyResult = (Var1 > Var2)    ' Returns True.

Var1 = 5: Var2 = Empty
MyResult = (Var1 > Var2)    ' Returns True.

Var1 = 0: Var2 = Empty
MyResult = (Var1 = Var2)    ' Returns True.

```


## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
