---
title: Rnd function (Visual Basic for Applications)
keywords: vblr6.chm1009008
f1_keywords:
- vblr6.chm1009008
ms.prod: office
ms.assetid: 57b9e8f9-6e3e-e68b-f5a4-c9c312b74426
ms.date: 12/13/2018
localization_priority: Normal
---


# Rnd function

Returns a **Single** containing a pseudo-random number.

## Syntax

**Rnd** [ (_Number_) ]

The optional _Number_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Single](../../Glossary/vbe-glossary.md#single-data-type) or any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).

## Return values

|If _Number_ is|Rnd generates|
|:-----|:-----|
|Less than zero|The same number every time, using _Number_ as the [seed](../../Glossary/vbe-glossary.md#seed).|
|Greater than zero|The next number in the pseudo-random sequence.|
|Equal to zero|The most recently generated number.|
|Not supplied|The next number in the pseudo-random sequence.|

## Remarks

The **Rnd** function returns a value less than 1 but greater than or equal to zero.

The value of _Number_ determines how **Rnd** generates a pseudo-random number:

- For any given initial seed, the same number sequence is generated because each successive call to the **Rnd** function uses the previous number as a seed for the next number in the sequence.

- Before calling **Rnd,** use the **Randomize** statement without an argument to initialize the random-number generator with a seed based on the system timer.

To produce random integers in a given range, use this formula:

```vb
Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

```

Here, _upperbound_ is the highest number in the range, and _lowerbound_ is the lowest number in the range.

> [!NOTE] 
> To repeat sequences of random numbers, call **Rnd** with a negative argument immediately before using **Randomize** with a numeric argument. Using **Randomize** with the same value for _Number_ does not repeat the previous sequence.

## Example

This example uses the **Rnd** function to generate a random integer value from 1 to 6.

```vb
Dim MyValue As Integer
MyValue = Int((6 * Rnd) + 1)    ' Generate random value between 1 and 6.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
