---
title: If...Then...Else statement (VBA)
keywords: vblr6.chm1008940
f1_keywords:
- vblr6.chm1008940
ms.prod: office
ms.assetid: 53514f63-ec20-27bf-2b61-5706540a4999
ms.date: 12/03/2018
localization_priority: Normal
---


# If...Then...Else statement

Conditionally executes a group of [statements](../../Glossary/vbe-glossary.md#statement), depending on the value of an [expression](../../Glossary/vbe-glossary.md#expression).

## Syntax

**If** _condition_ **Then** [ _statements_ ] [ **Else** _elsestatements_ ]

<br/>

Or, you can use the block form syntax:

**If** _condition_ **Then** <br/>
[ _statements_ ] <br/>
[ **ElseIf** _condition-n_ **Then** <br/>
[ _elseifstatements_ ]] <br/>
[ **Else** <br/>
[ _elsestatements_ ]] <br/>
**End If** <br/>

<br/>

The **If...Then...Else** statement syntax has these parts.

|Part|Description|
|:-----|:-----|
| _condition_|Required. One or more of the following two types of expressions: <br/><br/>A numeric expression or string expression that evaluates to **True** or **False**. If _condition_ is Null, _condition_ is treated as **False**.<br/><br/>An expression of the form **TypeOf** _objectname_ **Is** _objecttype_. The _objectname_ is any object reference, and _objecttype_ is any valid object type.  The expression is **True** if _objectname_ is of the object type specified by _objecttype_; otherwise, it is **False**.| 
| _statements_|Optional in block form; required in single-line form that has no **Else** clause. One or more statements separated by colons; executed if _condition_ is **True**.|
| _condition-n_|Optional. Same as  _condition_.|
| _elseifstatements_|Optional. One or more statements executed if associated  _condition-n_ is **True**.|
| _elsestatements_|Optional. One or more statements executed if no previous  _condition_ or _condition-n_ expression is **True**.|

## Remarks

You can use the single-line form (first syntax) for short, simple tests. However, the block form (second syntax) provides more structure and flexibility than the single-line form and is usually easier to read, maintain, and debug.

> [!NOTE] 
> With the single-line form, it is possible to have multiple statements executed as the result of an **If...Then** decision. All statements must be on the same line and separated by colons, as in the following statement:

```vb
If A > 10 Then A = A + 1 : B = B + A : C = C + B 

```


A block form **If** statement must be the first statement on a line. The **Else**, **ElseIf**, and **End If** parts of the statement can have only a line number or line label preceding them. The block **If** must end with an **End If** statement.

To determine whether or not a statement is a block **If**, examine what follows the **Then** keyword. If anything other than a comment appears after **Then** on the same line, the statement is treated as a single-line **If** statement.

The **Else** and **ElseIf** clauses are both optional. You can have as many **ElseIf** clauses as you want in a block **If**, but none can appear after an **Else** clause. Block **If** statements can be nested; that is, contained within one another.

When executing a block **If** (second syntax), _condition_ is tested. If _condition_ is **True**, the statements following **Then** are executed. If _condition_ is **False**, each **ElseIf** condition (if any) is evaluated in turn. When a **True** condition is found, the statements immediately following the associated **Then** are executed. If none of the **ElseIf** conditions are **True** (or if there are no **ElseIf** clauses), the statements following **Else** are executed. After executing the statements following **Then** or **Else**, execution continues with the statement following **End If**.

> [!TIP] 
> **Select Case** may be more useful when evaluating a single expression that has several possible actions. However, the **TypeOf** _objectname_ **Is** _objecttype_ clause can't be used with the **Select Case** statement.

> [!NOTE] 
> **TypeOf** cannot be used with hard data types such as Long, Integer, and so forth other than Object.


## Example

This example shows both the block and single-line forms of the **If...Then...Else** statement. It also illustrates the use of **If TypeOf...Then...Else**.


```vb
Dim Number, Digits, MyString 
Number = 53 ' Initialize variable. 
If Number < 10 Then 
 Digits = 1 
ElseIf Number < 100 Then 
' Condition evaluates to True so the next statement is executed. 
 Digits = 2 
Else 
 Digits = 3 
End If 
 
' Assign a value using the single-line form of syntax. 
If Digits = 1 Then MyString = "One" Else MyString = "More than one" 

```

<br/>

Use the **If TypeOf** construct to determine whether the Control passed into a procedure is a text box.

```vb
Sub ControlProcessor(MyControl As Control) 
 If TypeOf MyControl Is CommandButton Then 
 Debug.Print "You passed in a " & TypeName(MyControl) 
 ElseIf TypeOf MyControl Is CheckBox Then 
 Debug.Print "You passed in a " & TypeName(MyControl) 
 ElseIf TypeOf MyControl Is TextBox Then 
 Debug.Print "You passed in a " & TypeName(MyControl) 
 End If 
End Sub
```

## See also

- [If...Then...Else directive](ifthenelse-directive.md)
- [Using If...Then...Else statements](../../concepts/getting-started/using-ifthenelse-statements.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)
- [Visual Basic Editor (VBE) Glossary](../../Glossary/vbe-glossary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
