---
title: Select Case statement (VBA)
keywords: vblr6.chm1008810
f1_keywords:
- vblr6.chm1008810
ms.prod: office
ms.assetid: 8e885f14-c722-5217-705e-474516fa416b
ms.date: 12/03/2018
localization_priority: Normal
---


# Select Case statement

Executes one of several groups of [statements](../../Glossary/vbe-glossary.md#statement), depending on the value of an [expression](../../Glossary/vbe-glossary.md#expression).

## Syntax

**Select Case** _testexpression_ <br/>
 [ **Case** _expressionlist-n_ [ _statements-n_ ]] <br/>
 [ **Case Else** [ _elsestatements_ ]] <br/>
**End Select**

<br/>

The **Select Case** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _testexpression_|Required. Any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression).|
| _expressionlist-n_|Required if a **Case** appears.<br/><br/>Delimited list of one or more of the following forms: _expression_, _expression_**To**_expression_, **Is**_comparisonoperator_ _expression_.<br/><br/>The **To** [keyword](../../Glossary/vbe-glossary.md#keyword) specifies a range of values. If you use the **To** keyword, the smaller value must appear before **To**.<br/><br/>Use the **Is** keyword with [comparison operators](../../Glossary/vbe-glossary.md#comparison-operator) (except **Is** and **Like**) to specify a range of values. If not supplied, the **Is** keyword is automatically inserted.|
| _statements-n_|Optional. One or more statements executed if _testexpression_ matches any part of _expressionlist-n_. |
| _elsestatements_|Optional. One or more statements executed if _testexpression_ doesn't match any of the **Case** clause.|

## Remarks

If _testexpression_ matches any **Case** _expressionlist_ expression, the _statements_ following that **Case** clause are executed up to the next **Case** clause, or, for the last clause, up to **End Select**. Control then passes to the statement following **End Select**. If _testexpression_ matches an _expressionlist_ expression in more than one **Case** clause, only the statements following the first match are executed.

The **Case Else** clause is used to indicate the _elsestatements_ to be executed if no match is found between the _testexpression_ and an _expressionlist_ in any of the other **Case** selections. Although not required, it is a good idea to have a **Case Else** statement in your **Select Case** block to handle unforeseen _testexpression_ values. If no **Case** _expressionlist_ matches _testexpression_ and there is no **Case Else** statement, execution continues at the statement following **End Select**.

You can use multiple expressions or ranges in each **Case** clause. For example, the following line is valid:

```vb
Case 1 To 4, 7 To 9, 11, 13, Is > MaxNumber 

```

> [!NOTE] 
> The **Is** comparison operator is not the same as the **Is** keyword used in the **Select Case** statement.

You also can specify ranges and multiple expressions for character strings. In the following example, **Case** matches strings that are exactly equal to `everything`, strings that fall between `nuts` and `soup` in alphabetic order, and the current value of `TestItem`:

```vb
Case "everything", "nuts" To "soup", TestItem 

```

**Select Case** statements can be nested. Each nested **Select Case** statement must have a matching **[End Select](operator-summary.md)** statement.

## Example

This example uses the **Select Case** statement to evaluate the value of a variable. The second **Case** clause contains the value of the variable being evaluated, and therefore only the statement associated with it is executed.


```vb
Dim Number 
Number = 8    ' Initialize variable. 
Select Case Number    ' Evaluate Number. 
Case 1 To 5    ' Number between 1 and 5, inclusive. 
    Debug.Print "Between 1 and 5" 
' The following is the only Case clause that evaluates to True. 
Case 6, 7, 8    ' Number between 6 and 8. 
    Debug.Print "Between 6 and 8" 
Case 9 To 10    ' Number is 9 or 10. 
Debug.Print "Greater than 8" 
Case Else    ' Other values. 
    Debug.Print "Not between 1 and 10" 
End Select
```

## See also

- [Using Select Case statements](../../concepts/getting-started/using-select-case-statements.md)
- [Data types](data-type-summary.md)
- [Operators](operator-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
