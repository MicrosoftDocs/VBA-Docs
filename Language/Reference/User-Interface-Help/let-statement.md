---
title: Let statement (VBA)
keywords: vblr6.chm1008960
f1_keywords:
- vblr6.chm1008960
ms.prod: office
ms.assetid: da1ec875-3c6a-b66d-a85f-bbf33f9a307a
ms.date: 12/03/2018
localization_priority: Normal
---


# Let statement

Assigns the value of an [expression](../../Glossary/vbe-glossary.md#expression) to a [variable](../../Glossary/vbe-glossary.md#variable) or [property](../../Glossary/vbe-glossary.md#property).

## Syntax

[ **Let** ] _varname_ **=** _expression_

<br/>

The **Let** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Let**|Optional. Explicit use of the **Let** [keyword](../../Glossary/vbe-glossary.md#keyword) is a matter of style, but it is usually omitted.|
| _varname_|Required. Name of the variable or property; follows standard variable naming conventions.|
| _expression_|Required. Value assigned to the variable or property.|

## Remarks

A value expression can be assigned to a variable or property only if it is of a [data type](../../Glossary/vbe-glossary.md#data-type) that is compatible with the variable. You can't assign [string expressions](../../Glossary/vbe-glossary.md#string-expression) to numeric variables, and you can't assign [numeric expressions](../../Glossary/vbe-glossary.md#numeric-expression) to string variables. If you do, an error occurs at [compile time](../../Glossary/vbe-glossary.md#compile-time).

[Variant](../../Glossary/vbe-glossary.md#variant-data-type) variables can be assigned to either string or numeric expressions. However, the reverse is not always true. Any **Variant** except a [Null](../../Glossary/vbe-glossary.md#null) can be assigned to a string variable, but only a **Variant** whose value can be interpreted as a number can be assigned to a numeric variable. Use the **IsNumeric** function to determine if the **Variant** can be converted to a number.

Assigning an expression of one [numeric type](../../Glossary/vbe-glossary.md#numeric-type) to a variable of a different numeric type coerces the value of the expression into the numeric type of the resulting variable.

**Let** statements can be used to assign one record variable to another only when both variables are of the same [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type). Use the **[LSet](lset-statement.md)** statement to assign record variables of different user-defined types. Use the **[Set](set-statement.md)** statement to assign object references to variables.

## Example

This example assigns the values of expressions to variables by using the explicit **Let** statement.

```vb
Dim MyStr, MyInt 
' The following variable assignments use the Let statement. 
Let MyStr = "Hello World" 
Let MyInt = 5 

```

<br/>

The following are the same assignments without the **Let** statement.

```vb
Dim MyStr, MyInt 
MyStr = "Hello World" 
MyInt = 5 

```


## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
