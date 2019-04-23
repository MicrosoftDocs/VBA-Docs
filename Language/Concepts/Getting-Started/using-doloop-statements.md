---
title: Using Do...Loop statements (VBA)
keywords: vbcn6.chm1076681
f1_keywords:
- vbcn6.chm1076681
ms.prod: office
ms.assetid: aa3322b6-80a6-d3c6-86b7-4ea6151f0616
ms.date: 12/26/2018
localization_priority: Normal
---


# Using Do...Loop statements

You can use **[Do...Loop](../../reference/user-interface-help/doloop-statement.md)** statements to run a block of [statements](../../Glossary/vbe-glossary.md#statement) an indefinite number of times. The statements are repeated either while a condition is **True** or until a condition becomes **True**.


## Repeating statements while a condition is True

There are two ways to use the **While** [keyword](../../Glossary/vbe-glossary.md#keyword) to check a condition in a **Do...Loop** statement. You can check the condition before you enter the loop, or you can check it after the loop has run at least once.

In the following `ChkFirstWhile` procedure, you check the condition before you enter the loop. If `myNum` is set to 9 instead of 20, the statements inside the loop will never run. In the `ChkLastWhile` procedure, the statements inside the loop run only once before the condition becomes **False**.

```vb
Sub ChkFirstWhile() 
    counter = 0 
    myNum = 20 
    Do While myNum > 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub 
 
Sub ChkLastWhile() 
    counter = 0 
    myNum = 9 
    Do 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop While myNum > 10 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```


## Repeating statements until a condition becomes True

There are two ways to use the **Until** keyword to check a condition in a **Do...Loop** statement. You can check the condition before you enter the loop (as shown in the `ChkFirstUntil` procedure), or you can check it after the loop has run at least once (as shown in the `ChkLastUntil` procedure). Looping continues while the condition remains **False**.


```vb
Sub ChkFirstUntil() 
    counter = 0 
    myNum = 20 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub 
 
Sub ChkLastUntil() 
    counter = 0 
    myNum = 1 
    Do 
        myNum = myNum + 1 
        counter = counter + 1 
    Loop Until myNum = 10 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```


## Exiting a Do...Loop statement from inside the loop

You can exit a **Do...Loop** by using the **[Exit Do](../../reference/user-interface-help/exit-statement.md)** statement. For example, to exit an endless loop, use the **Exit Do** statement in the **True** statement block of either an **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statement or a **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement. If the condition is **False**, the loop will run as usual.

In the following example `myNum` is assigned a value that creates an endless loop. The **If...Then...Else** statement checks for this condition, and then exits, preventing endless looping.

```vb
Sub ExitExample() 
    counter = 0 
    myNum = 9 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
        If myNum < 10 Then Exit Do 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```

> [!NOTE] 
> To stop an endless loop, press ESC or CTRL+BREAK.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
