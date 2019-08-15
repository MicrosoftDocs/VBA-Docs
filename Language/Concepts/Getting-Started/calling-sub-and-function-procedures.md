---
title: Calling Sub and Function procedures (VBA)
keywords: vbcn6.chm1076673
f1_keywords:
- vbcn6.chm1076673
ms.prod: office
ms.assetid: 17a9dec1-d8f2-584c-324f-164b4f7b156f
ms.date: 08/14/2019
localization_priority: Priority
---


# Calling Sub and Function procedures

To call a **[Sub](../../reference/user-interface-help/sub-statement.md)** procedure from another [procedure](../../Glossary/vbe-glossary.md#procedure), type the name of the procedure and include values for any required [arguments](../../Glossary/vbe-glossary.md#argument). The **[Call](../../reference/user-interface-help/call-statement.md)** statement is not required, but if you use it, you must enclose any arguments in parentheses.

You can use a **Sub** procedure to organize other procedures so they are easier to understand and debug. In the following example, the **Sub** procedure `Main` calls the **Sub** procedure `MultiBeep`, passing the value 56 for its argument. 

After `MultiBeep` runs, control returns to `Main`, and `Main` calls the **Sub** procedure `Message`. `Message` displays a message box; when the user clicks **OK**, control returns to `Main`, and `Main` finishes.

```vb
Sub Main() 
 MultiBeep 56 
 Message 
End Sub 
 
Sub MultiBeep(numbeeps) 
 For counter = 1 To numbeeps 
 Beep 
 Next counter 
End Sub 
 
Sub Message() 
 MsgBox "Time to take a break!" 
End Sub
```

[!include[Add-ins note](~/includes/addinsnote.md)]

## Call Sub procedures with more than one argument

The following example shows two ways to call a **Sub** procedure with more than one argument. The second time it is called, parentheses are required around the arguments because the **Call** statement is used.

```vb
Sub Main() 
 HouseCalc 99800, 43100 
 Call HouseCalc(380950, 49500) 
End Sub 
 
Sub HouseCalc(price As Single, wage As Single) 
 If 2.5 * wage <= 0.8 * price Then 
 MsgBox "You cannot afford this house." 
 Else 
 MsgBox "This house is affordable." 
 End If 
End Sub
```


## Use parentheses when calling function procedures

To use the return value of a function, assign the function to a [variable](../../Glossary/vbe-glossary.md#variable) and enclose the arguments in parentheses, as shown in the following example.

```vb
Answer3 = MsgBox("Are you happy with your salary?", 4, "Question 3") 

```

If you are not interested in the return value of a function, you can call a function the same way you call a **Sub** procedure. Omit the parentheses, list the arguments, and do not assign the function to a variable, as shown in the following example.

```vb
MsgBox "Task Completed!", 0, "Task Box" 

```

If you include parentheses in the preceding example, the statement causes a syntax error.


## Pass named arguments

A statement in a **Sub** or **[Function](../../reference/user-interface-help/function-statement.md)** procedure can pass values to called procedures by using [named arguments](../../Glossary/vbe-glossary.md#named-argument). You can list named arguments in any order. A named argument consists of the name of the argument followed by a colon and an equal sign (**:=**), and the value assigned to the argument.

The following example calls the **MsgBox** function by using named arguments with no return value.

```vb
MsgBox Title:="Task Box", Prompt:="Task Completed!" 

```

The following example calls the **MsgBox** function by using named arguments. The return value is assigned to the variable.


```vb
answer3 = MsgBox(Title:="Question 3", _ 
Prompt:="Are you happy with your salary?", Buttons:=4) 

```

## See also

- [Using parentheses in code](using-parentheses-in-code.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
