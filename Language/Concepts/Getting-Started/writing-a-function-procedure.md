---
title: Writing a Function procedure (VBA)
keywords: vbcn6.chm1076690
f1_keywords:
- vbcn6.chm1076690
ms.prod: office
ms.assetid: 80e2ad00-a12f-2f40-3cb8-9878a595dde3
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing a Function procedure

A **[Function](../../reference/user-interface-help/function-statement.md)** procedure is a series of Visual Basic [statements](../../Glossary/vbe-glossary.md#statement) enclosed by the **Function** and **[End Function](../../reference/user-interface-help/end-statement.md)** statements. A **Function** procedure is similar to a **[Sub](../../reference/user-interface-help/sub-statement.md)** procedure, but a function can also return a value. 

A **Function** procedure can take [arguments](../../Glossary/vbe-glossary.md#argument), such as [constants](../../Glossary/vbe-glossary.md#constant), [variables](../../Glossary/vbe-glossary.md#variable), or [expressions](../../Glossary/vbe-glossary.md#expression) that are passed to it by a calling procedure. If a **Function** procedure has no arguments, its **Function** statement must include an empty set of parentheses. A function returns a value by assigning a value to its name in one or more statements of the procedure.

In the following example, the **Celsius** function calculates degrees Celsius from degrees Fahrenheit. When the function is called from the **Main** procedure, a variable containing the argument value is passed to the function. The result of the calculation is returned to the calling procedure and displayed in a message box.

```vb
Sub Main() 
 temp = Application.InputBox(Prompt:= _ 
 "Please enter the temperature in degrees F.", Type:=1) 
 MsgBox "The temperature is " & Celsius(temp) & " degrees C." 
End Sub 
 
Function Celsius(fDegrees) 
 Celsius = (fDegrees - 32) * 5 / 9 
End Function
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
