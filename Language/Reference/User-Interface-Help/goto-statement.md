---
title: GoTo statement (VBA)
keywords: vblr6.chm1008935
f1_keywords:
- vblr6.chm1008935
ms.prod: office
ms.assetid: 0fa45435-77cf-91f5-ade4-86ac0eb1a083
ms.date: 12/03/2018
localization_priority: Normal
---


# GoTo statement

Branches unconditionally to a specified line within a [procedure](../../Glossary/vbe-glossary.md#procedure).

## Syntax

**GoTo** _line_

The required _line_ [argument](../../Glossary/vbe-glossary.md#argument) can be any [line label](../../Glossary/vbe-glossary.md#line-label) or [line number](../../Glossary/vbe-glossary.md#line-number).

## Remarks

**GoTo** can branch only to lines within the procedure where it appears.

> [!NOTE] 
> Too many **GoTo** statements can make code difficult to read and debug. Use structured control [statements](../../Glossary/vbe-glossary.md#statement) (**[Do...Loop](doloop-statement.md)**, **[For...Next](fornext-statement.md)**, **[If...Then...Else](ifthenelse-statement.md)**, **[Select Case](select-case-statement.md)**) whenever possible.

## Example

This example uses the **GoTo** statement to branch to line labels within a procedure.


```vb
Sub GotoStatementDemo() 
Dim Number, MyString 
 Number = 1 ' Initialize variable. 
 ' Evaluate Number and branch to appropriate label. 
 If Number = 1 Then GoTo Line1 Else GoTo Line2 
 
Line1: 
 MyString = "Number equals 1" 
 GoTo LastLine ' Go to LastLine. 
Line2: 
 ' The following statement never gets executed. 
 MyString = "Number equals 2" 
LastLine: 
 Debug.Print MyString ' Print "Number equals 1" in 
 ' the Immediate window. 
End Sub
```


## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
