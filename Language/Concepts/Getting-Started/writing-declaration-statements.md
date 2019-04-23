---
title: Writing declaration statements (VBA)
keywords: vbcn6.chm1076693
f1_keywords:
- vbcn6.chm1076693
ms.prod: office
ms.assetid: 9aaee08c-09d3-b70b-0d8f-9ca949fbd04a
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing declaration statements

You use declaration statements to name and define [procedures](../../Glossary/vbe-glossary.md#procedure), [variables](../../Glossary/vbe-glossary.md#variable), [arrays](../../Glossary/vbe-glossary.md#array), and [constants](../../Glossary/vbe-glossary.md#constant). When you declare a procedure, variable, or constant, you also define its [scope](../../Glossary/vbe-glossary.md#scope), depending on where you place the declaration and what [keywords](../../Glossary/vbe-glossary.md#keyword) you use to declare it.

The following example contains three declarations.

```vb
Sub ApplyFormat() 
    Const limit As Integer = 33 
    Dim myCell As Range 
    ' More statements 
End Sub
```

The **Sub** statement (with matching **End Sub** statement) declares a procedure named `ApplyFormat`. All the statements enclosed by the **Sub** and **End Sub** statements are executed whenever the `ApplyFormat` procedure is called or run.

The **Const** statement declares the constant `limit` specifying the **Integer** data type and a value of 33.

The **Dim** statement declares the `myCell` variable. The data type is an object, in this case, a Microsoft Excel **Range** object. You can declare a variable to be any object that is exposed in the application that you are using. **Dim** statements are one type of statement used to declare variables. Other keywords used in declarations are **ReDim**, **Static**, **Public**, **Private**, and **Const**.

## See also

- [Statements](../../reference/statements.md)
- [Writing a Sub procedure](writing-a-sub-procedure.md)
- [Declaring constants](declaring-constants.md)
- [Declaring variables](declaring-variables.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]