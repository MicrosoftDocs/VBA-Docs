---
title: Understanding Visual Basic syntax (VBA)
keywords: vbcn6.chm1076679
f1_keywords:
- vbcn6.chm1076679
ms.prod: office
ms.assetid: 8b6f4203-f82e-5f2f-ad1e-1ad90d088700
ms.date: 12/26/2018
localization_priority: Normal
---


# Understanding Visual Basic syntax

The syntax in a Visual Basic Help topic for a [method](../../Glossary/vbe-glossary.md#method), [function](../../Glossary/vbe-glossary.md#function-procedure), or [statement](../../Glossary/vbe-glossary.md#statement) shows all the elements necessary to use the method, function, or statement correctly. The examples in this topic explain how to interpret the most common syntax elements.

## Activate method syntax

_object_.**Activate**

In the **Activate** method syntax, the italic word "object" is a placeholder for information you supply—in this case, code that returns an [object](../../Glossary/vbe-glossary.md#object). Words that are bold should be typed exactly as they appear. For example, the following [procedure](../../Glossary/vbe-glossary.md#procedure) activates the second window in the active document.

```vb
Sub MakeActive() 
    Windows(2).Activate 
End Sub
```

## MsgBox function syntax

**MsgBox** (_prompt_, [ _buttons_, ] [ _title_, ] [ _helpfile_, _context_ ])

In the **[MsgBox](../../reference/user-interface-help/msgbox-function.md)** function syntax, the italic words are [named arguments](../../Glossary/vbe-glossary.md#named-argument) of the function. [Arguments](../../Glossary/vbe-glossary.md#argument) enclosed in brackets are optional. (Do not type the brackets in your Visual Basic code.) For the **MsgBox** function, the only argument you must provide is the text for the prompt.

Arguments for functions and methods can be specified in code either by position or by name. To specify arguments by position, follow the order presented in the syntax, separating each argument with a comma, for example:

```vb
MsgBox "Your answer is correct!",0,"Answer Box" 

```

To specify an argument by name, use the argument name followed by a colon and an equal sign (**:=**), and the argument's value. You can specify named arguments in any order, for example:

```vb
MsgBox Title:="Answer Box", Prompt:="Your answer is correct!" 

```

The syntax for functions and some methods shows the arguments enclosed in parentheses. These functions and methods return values, so you must enclose the arguments in parentheses to assign the value to a variable. If you ignore the return value or if you don't pass arguments at all, don't include the parentheses. Methods that don't return values do not need their arguments enclosed in parentheses. These guidelines apply whether you are using positional arguments or named arguments.

In the following example, the return value from the **MsgBox** function is a number indicating the selected button that is stored in the variable `myVar`. Because the return value is used, parentheses are required. Another message box then displays the value of the variable.

```vb
Sub Question() 
    myVar = MsgBox(Prompt:="I enjoy my job.", _ 
        Title:="Answer Box", Buttons:="4") 
    MsgBox myVar 
End Sub
```

## Option Compare statement syntax

**Option Compare** { **Binary** | **Text** | **Database** }

In the **[Option Compare](../../reference/user-interface-help/option-compare-statement.md)** statement syntax, the braces and vertical bar indicate a mandatory choice between three items. (Do not type the braces in the Visual Basic statement). For example, the following statement specifies that within the [module](../../Glossary/vbe-glossary.md#module), strings will be compared in a [sort order](../../Glossary/vbe-glossary.md#sort-order) that is not case-sensitive.

```vb
Option Compare Text 

```

## Dim statement syntax

**Dim** _varname_ [([ _subscripts_ ])] [ **As** _type_, ] [ _varname_ [([ _subscripts_ ])] [ **As** _type_ ]] **. . .**

In the **[Dim](../../reference/user-interface-help/dim-statement.md)** statement syntax, the word **Dim** is a required [keyword](../../Glossary/vbe-glossary.md#keyword). The only required element is _varname_ (the variable name). 

For example, the following statement creates three variables: `myVar`, `nextVar`, and `thirdVar`. These are automatically declared as **Variant** variables.

```vb
Dim myVar, nextVar, thirdVar 

```

The following example declares a variable as a **String**. Including a [data type](../../Glossary/vbe-glossary.md#data-type) saves memory and can help you find errors in your code.

```vb
Dim myAnswer As String 

```

To declare several variables in one statement, include the data type for each variable. Variables declared without a data type are automatically declared as **Variant**.

```vb
Dim x As Integer, y As Integer, z As Integer 

```

In the following statement, `x` and `y` are assigned the **Variant** data type. Only `z` is assigned the **Integer** data type.

```vb
Dim x, y, z As Integer 

```

The shorthand to declare `x` and `y` as Integer in the statement above is:

```vb
Dim x%, y%, z as Integer

```
The shorthand for the types is: % -integer; & -long; @ -currency; # -double; ! -single; $ -string


If you are declaring an [array](../../Glossary/vbe-glossary.md#array) variable, you must include parentheses. The subscripts are optional. The following statement dimensions a dynamic array, `myArray`.

```vb
Dim myArray() 

```

## See also

- [Data types](../../reference/user-interface-help/data-type-summary.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
