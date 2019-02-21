---
title: Understanding parameter arrays (VBA)
keywords: vbcn6.chm1076759
f1_keywords:
- vbcn6.chm1076759
ms.prod: office
ms.assetid: 42438a68-37a8-85d0-6404-1df4266fe33d
ms.date: 12/26/2018
localization_priority: Normal
---


# Understanding parameter arrays

A [parameter](../../Glossary/vbe-glossary.md#parameter) [array](../../Glossary/vbe-glossary.md#array) can be used to pass an array of [arguments](../../Glossary/vbe-glossary.md#argument) to a [procedure](../../Glossary/vbe-glossary.md#procedure). You don't have to know the number of elements in the array when you define the procedure.

You use the **ParamArray** keyword to denote a parameter array. The array must be declared as an array of type **Variant**, and it must be the last argument in the procedure definition.

The following example shows how you might define a procedure with a parameter array.

```vb
Sub AnyNumberArgs(strName As String, ParamArray intScores() As Variant) 
 Dim intI As Integer 
 
 Debug.Print strName; " Scores" 
 ' Use UBound function to determine upper limit of array. 
 For intI = 0 To UBound(intScores()) 
 Debug.Print " "; intScores(intI) 
 Next intI 
End Sub
```

<br/>

The following examples show how you can call this procedure.

```vb
AnyNumberArgs "Jamie", 10, 26, 32, 15, 22, 24, 16 
 
AnyNumberArgs "Kelly", "High", "Low", "Average", "High" 
```

## See also

- [Visual Basic reference](../../reference/user-interface-help/visual-basic-language-reference.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]