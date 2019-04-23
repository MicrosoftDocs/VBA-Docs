---
title: Randomize statement (VBA)
keywords: vblr6.chm1008998
f1_keywords:
- vblr6.chm1008998
ms.prod: office
ms.assetid: b09ed4eb-1e05-c904-7cd5-482fea785ce6
ms.date: 12/03/2018
localization_priority: Normal
---


# Randomize statement

Initializes the random-number generator.

## Syntax

**Randomize** [ _number_ ]

The optional _number_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) or any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).

## Remarks

**Randomize** uses _number_ to initialize the **[Rnd](rnd-function.md)** function's random-number generator, giving it a new [seed](../../Glossary/vbe-glossary.md#seed) value. If you omit _number_, the value returned by the system timer is used as the new seed value.

If **Randomize** is not used, the **Rnd** function (with no arguments) uses the same number as a seed the first time it is called, and thereafter uses the last generated number as a seed value.

> [!NOTE] 
> To repeat sequences of random numbers, call **Rnd** with a negative argument immediately before using **Randomize** with a numeric argument. Using **Randomize** with the same value for _number_ does not repeat the previous sequence.


## Example

This example uses the **Randomize** statement to initialize the random-number generator. Because the number argument has been omitted, **Randomize** uses the return value from the **Timer** function as the new seed value.


```vb
Dim MyValue 
Randomize ' Initialize random-number generator. 
 
MyValue = Int((6 * Rnd) + 1) ' Generate random value between 1 and 6. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
