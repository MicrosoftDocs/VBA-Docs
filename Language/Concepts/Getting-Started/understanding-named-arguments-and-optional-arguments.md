---
title: Understanding named arguments and optional arguments (VBA)
keywords: vbcn6.chm1076747
f1_keywords:
- vbcn6.chm1076747
ms.prod: office
ms.assetid: 207fa305-56cf-4b44-d23e-dcc3b104b2dd
ms.date: 12/21/2018
localization_priority: Normal
---


# Understanding named arguments and optional arguments

When you call a **[Sub](../../reference/user-interface-help/sub-statement.md)** or **[Function](../../reference/user-interface-help/function-statement.md)** procedure, you can supply [arguments](../../Glossary/vbe-glossary.md#argument) positionally, in the order that they appear in the procedure's definition, or you can supply the arguments by name without regard to position.

For example, the following **Sub** procedure takes three arguments.

```vb
Sub PassArgs(strName As String, intAge As Integer, dteBirth As Date) 
 Debug.Print strName, intAge, dteBirth 
End Sub
```

<br/>

You can call this procedure by supplying its arguments in the correct position, each delimited by a comma, as shown in the following example.

```vb
PassArgs "Mary", 29, #2-21-69# 

```

<br/>

You can also call this procedure by supplying [named arguments](../../Glossary/vbe-glossary.md#named-argument), delimiting each with a comma.

```vb
PassArgs intAge:=29, dteBirth:=#2/21/69#, strName:="Mary" 

```

<br/>

A named argument consists of an argument name followed by a colon and an equal sign (**:=**), followed by the argument value.

Named arguments are especially useful when you are calling a procedure that has optional arguments. If you use named arguments, you don't have to include commas to denote missing positional arguments. Using named arguments makes it easier to keep track of which arguments you passed and which you omitted.

Optional arguments are preceded by the **Optional** [keyword](../../Glossary/vbe-glossary.md#keyword) in the procedure definition. You can also specify a default value for the optional argument in the procedure definition. For example:

```vb
Sub OptionalArgs(strState As String, Optional strCountry As String = "USA") 
. . . 
End Sub
```

<br/>

When you call a procedure with an optional argument, you can choose whether or not to specify the optional argument. If you don't specify the optional argument, the default value, if any, is used. If no default value is specified, the argument would be for any variable of the specified type.

The following procedure includes two optional arguments, the  `varRegion` and `varCountry` variables. The **[IsMissing](../../reference/user-interface-help/ismissing-function.md)** function determines whether an optional Variant argument has been passed to the procedure.

```vb
Sub OptionalArgs(strState As String, Optional varRegion As Variant, _ 
Optional varCountry As Variant = "USA") 
 If IsMissing(varRegion) And IsMissing(varCountry) Then 
 Debug.Print strState 
 ElseIf IsMissing(varCountry) Then 
 Debug.Print strState, varRegion 
 ElseIf IsMissing(varRegion) Then 
 Debug.Print strState, varCountry 
 Else 
 Debug.Print strState, varRegion, varCountry 
 End If 
End Sub
```

<br/>

You can call this procedure by using named arguments as shown in the following example.

```vb
OptionalArgs varCountry:="USA", strState:="MD" 
 
OptionalArgs strState:= "MD", varRegion:=5 

```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
