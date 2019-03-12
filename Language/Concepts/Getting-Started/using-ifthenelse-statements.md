---
title: Using If...Then...Else statements (VBA)
keywords: vbcn6.chm1076684
f1_keywords:
- vbcn6.chm1076684
ms.prod: office
ms.assetid: 5b138720-aad6-df90-305e-72adb644d7dd
ms.date: 12/26/2018
localization_priority: Priority
---


# Using If...Then...Else statements

You can use the **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statement to run a specific [statement](../../Glossary/vbe-glossary.md#statement) or a block of statements, depending on the value of a condition. **If...Then...Else** statements can be nested to as many levels as you need. 

However, for readability, you may want to use a **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement rather than multiple levels of nested **If...Then...Else** statements.


## Running statements if a condition is True

To run only one statement when a condition is **True**, use the single-line syntax of the **If...Then...Else** statement. The following example shows the single-line syntax, omitting the **Else** [keyword](../../Glossary/vbe-glossary.md#keyword).


```vb
Sub FixDate() 
 myDate = #2/13/95# 
 If myDate < Now Then myDate = Now 
End Sub
```

To run more than one line of code, you must use the multiple-line syntax. This syntax includes the **[End If](../../reference/user-interface-help/end-statement.md)** statement, as shown in the following example.

```vb
Sub AlertUser(value as Long) 
 If value = 0 Then 
 AlertLabel.ForeColor = "Red" 
 AlertLabel.Font.Bold = True 
 AlertLabel.Font.Italic = True 
 End If 
End Sub
```


## Running certain statements if a condition is True and running others if it's False

Use an **If...Then...Else** statement to define two blocks of executable statements: one block runs if the condition is **True**, and the other block runs if the condition is **False**.


```vb
Sub AlertUser(value as Long) 
 If value = 0 Then 
 AlertLabel.ForeColor = vbRed 
 AlertLabel.Font.Bold = True 
 AlertLabel.Font.Italic = True 
 Else 
 AlertLabel.Forecolor = vbBlack 
 AlertLabel.Font.Bold = False 
 AlertLabel.Font.Italic = False 
 End If 
End Sub
```


## Testing a second condition if the first condition is False

You can add **ElseIf** statements to an **If...Then...Else** statement to test a second condition if the first condition is **False**. For example, the following function procedure computes a bonus based on job classification. The statement following the **Else** statement runs if the conditions in all of the **If** and **ElseIf** statements are **False**.


```vb
Function Bonus(performance, salary) 
 If performance = 1 Then 
 Bonus = salary * 0.1 
 ElseIf performance = 2 Then 
 Bonus = salary * 0.09 
 ElseIf performance = 3 Then 
 Bonus = salary * 0.07 
 Else 
 Bonus = 0 
 End If 
End Function
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
