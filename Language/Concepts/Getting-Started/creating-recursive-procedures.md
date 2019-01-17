---
title: Creating recursive procedures (VBA)
keywords: vbcn6.chm1010995
f1_keywords:
- vbcn6.chm1010995
ms.prod: office
ms.assetid: 5458afe3-63ec-d2c2-8278-f6b5ce1734d3
ms.date: 12/21/2018
localization_priority: Normal
---


# Creating recursive procedures

[Procedures](../../Glossary/vbe-glossary.md#procedure) have a limited amount of space for [variables](../../Glossary/vbe-glossary.md#variable). Each time a procedure calls itself, more of that space is used. A procedure that calls itself is a recursive procedure. A recursive procedure that continuously calls itself eventually causes an error. For example:

```vb
Function RunOut(Maximum) 
 RunOut = RunOut(Maximum) 
End Function
```

This error may be less obvious when two procedures call each other indefinitely, or when some condition that limits the recursion is never met. Recursion does have its uses. For example, the following procedure uses a recursive function to calculate factorials.

```vb
Function Factorial (N) 
 If N <= 1 Then ' Reached end of recursive calls. 
 Factorial = 1 ' (N = 0) so climb back out of calls. 
 Else ' Call Factorial again if N > 0. 
 Factorial = Factorial(N - 1) * N 
 End If 
End Function
```

You should test your recursive procedure to make sure it does not call itself so many times that you run out of memory. If you get an error, make sure your procedure is not calling itself indefinitely. After that, try to conserve memory by:

- Eliminating unnecessary variables.
- Using [data types](../../reference/user-interface-help/data-type-summary.md) other than **Variant**.
- Re-evaluating the logic of the procedure. You can often substitute nested loops for recursion.
    

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]