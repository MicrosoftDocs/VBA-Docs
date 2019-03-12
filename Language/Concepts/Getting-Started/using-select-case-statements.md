---
title: Using Select Case statements (VBA)
keywords: vbcn6.chm1076686
f1_keywords:
- vbcn6.chm1076686
ms.prod: office
ms.assetid: 0573a361-84d6-549f-8c51-5bc0fe17d156
ms.date: 12/26/2018
localization_priority: Normal
---


# Using Select Case statements

Use the **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement as an alternative to using **ElseIf** in **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statements when comparing one [expression](../../Glossary/vbe-glossary.md#expression) to several different values. While **If...Then...Else** statements can evaluate a different expression for each **ElseIf** statement, the **Select Case** statement evaluates an expression only once, at the top of the control structure.

In the following example, the **Select Case** statement evaluates the argument that is passed to the procedure. Note that each **Case** statement can contain more than one value, a range of values, or a combination of values and [comparison operators](../../Glossary/vbe-glossary.md#comparison-operator). The optional **Case Else** statement runs if the **Select Case** statement doesn't match a value in any of the **Case** statements.

```vb
Function Bonus(performance, salary) 
  Select Case performance 
    Case 1 
      Bonus = salary * 0.1 
    Case 2, 3 
      Bonus = salary * 0.09 
    Case 4 To 6 
      Bonus = salary * 0.07 
    Case Is > 8 
      Bonus = 100 
    Case Else 
      Bonus = 0 
  End Select 
End Function 
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
