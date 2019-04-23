---
title: Writing Visual Basic statements (VBA)
keywords: vbcn6.chm1076695
f1_keywords:
- vbcn6.chm1076695
ms.prod: office
ms.assetid: a2d35638-995b-1a6f-2975-8deacddf93de
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing Visual Basic statements

A [statement](../../Glossary/vbe-glossary.md#statement) in Visual Basic is a complete instruction. It can contain [keywords](../../Glossary/vbe-glossary.md#keyword), operators, [variables](../../Glossary/vbe-glossary.md#variable), [constants](../../Glossary/vbe-glossary.md#constant), and [expressions](../../Glossary/vbe-glossary.md#expression). Each statement belongs to one of the following three categories:

- [Declaration statements](writing-declaration-statements.md), which name a variable, constant, or procedure and can also specify a data type. 
    
- [Assignment statements](writing-assignment-statements.md), which assign a value or expression to a variable or constant.
    
- [Executable statements](writing-executable-statements.md), which initiate actions. These statements can execute a method or function, and they can loop or branch through blocks of code. Executable statements often contain mathematical or conditional operators.
    
## Continue a statement over multiple lines

A statement usually fits on one line, but you can continue a statement onto the next line by using a [line-continuation character](../../Glossary/vbe-glossary.md#line-continuation-character). In the following example, the **MsgBox** executable statement is continued over three lines:

```vb
Sub DemoBox() 'This procedure declares a string variable, 
 ' assigns it the value Claudia, and then displays 
 ' a concatenated message. 
 Dim myVar As String 
 myVar = "John" 
 MsgBox Prompt:="Hello " & myVar, _ 
 Title:="Greeting Box", _ 
 Buttons:=vbExclamation 
End Sub
```


## Add comments

Comments can explain a procedure or a particular instruction to anyone reading your code. Visual Basic ignores comments when it runs your procedures. Comment lines begin with an apostrophe (**'**) or with **Rem** followed by a space, and can be added anywhere in a procedure. To add a comment to the same line as a statement, insert an apostrophe after the statement, followed by the comment. By default, comments are displayed as green text.


## Check syntax errors

If you press ENTER after typing a line of code and the line is displayed in red (an error message may display as well), you must find out what's wrong with your statement, and then correct it.


## See also

- [Statements](../../reference/statements.md)
- [Document conventions](document-conventions-visual-basic-for-applications.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
