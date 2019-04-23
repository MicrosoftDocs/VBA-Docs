---
title: Writing a Sub procedure (VBA)
keywords: vbcn6.chm1076691
f1_keywords:
- vbcn6.chm1076691
ms.prod: office
ms.assetid: 5c9b6ff6-d8a4-7a4f-577f-da9f3257bb44
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing a Sub procedure

A **[Sub](../../reference/user-interface-help/sub-statement.md)** procedure is a series of Visual Basic [statements](../../Glossary/vbe-glossary.md#statement) enclosed by the **Sub** and **[End Sub](../../reference/user-interface-help/end-statement.md)** statements that performs actions but doesn't return a value. A **Sub** procedure can take arguments, such as [constants](../../Glossary/vbe-glossary.md#constant), [variables](../../Glossary/vbe-glossary.md#variable), or [expressions](../../Glossary/vbe-glossary.md#expression) that are passed by a calling procedure. If a **Sub** procedure has no arguments, the **Sub** statement must include an empty set of parentheses.

The following **Sub** procedure has comments explaining each line.

```vb
' Declares a procedure named GetInfo 
' This Sub procedure takes no arguments 
Sub GetInfo() 
' Declares a string variable named answer 
Dim answer As String 
' Assigns the return value of the InputBox function to answer 
answer = InputBox(Prompt:="What is your name?") 
 ' Conditional If...Then...Else statement 
 If answer = Empty Then 
 ' Calls the MsgBox function 
 MsgBox Prompt:="You did not enter a name." 
 Else 
 ' MsgBox function concatenated with the variable answer 
 MsgBox Prompt:="Your name is " & answer 
 ' Ends the If...Then...Else statement 
 End If 
' Ends the Sub procedure 
End Sub
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
