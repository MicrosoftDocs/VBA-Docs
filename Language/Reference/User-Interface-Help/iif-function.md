---
title: IIf function (Visual Basic for Applications)
keywords: vblr6.chm1012957
f1_keywords:
- vblr6.chm1012957
ms.prod: office
ms.assetid: a31d9f49-1f5a-324b-77a2-276eb573552a
ms.date: 12/13/2018
localization_priority: Normal
---


# IIf function

Returns one of two parts, depending on the evaluation of an [expression](../../Glossary/vbe-glossary.md#expression).

## Syntax

**IIf**(_expr_, _truepart_, _falsepart_)

<br/>

The **IIf** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_expr_|Required. Expression that you want to evaluate.|
|_truepart_|Required. Value or expression returned if _expr_ is **True**.|
|_falsepart_|Required. Value or expression returned if _expr_ is **False**.|

## Remarks

**IIf** always evaluates both _truepart_ and _falsepart_, even though it returns only one of them. Because of this, you should watch for undesirable side effects. For example, if evaluating _falsepart_ results in a division by zero error, an error occurs even if _expr_ is **True**.

## Example

This example uses the **IIf** function to evaluate the `TestMe` parameter of the `CheckIt` procedure and returns the word "Large" if the amount is greater than 1000; otherwise, it returns the word "Small".

```vb
Function CheckIt (TestMe As Integer)
    CheckIt = IIf(TestMe > 1000, "Large", "Small")
End Function
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
