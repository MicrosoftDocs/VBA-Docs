---
title: Using For...Next statements (VBA)
keywords: vbcn6.chm1076682
f1_keywords:
- vbcn6.chm1076682
ms.prod: office
ms.assetid: fe6e66a7-a9d3-d363-65c5-00d35bb407bd
ms.date: 12/26/2018
localization_priority: Normal
---


# Using For...Next statements

You can use **[For...Next](../../reference/user-interface-help/fornext-statement.md)** statements to repeat a block of [statements](../../Glossary/vbe-glossary.md#statement) a specific number of times. **For** loops use a counter [variable](../../Glossary/vbe-glossary.md#variable) whose value is increased or decreased with each repetition of the loop.

The following [procedure](../../Glossary/vbe-glossary.md#procedure) makes the computer beep 50 times. The **For** statement specifies the counter variable and its start and end values. The **Next** statement increments the counter variable by 1.

```vb
Sub Beeps() 
    For x = 1 To 50 
        Beep 
    Next x 
End Sub
```

Using the **Step** [keyword](../../Glossary/vbe-glossary.md#keyword), you can increase or decrease the counter variable by the value you specify. In the following example, the counter variable `j` is incremented by 2 each time the loop repeats. When the loop is finished, `total` is the sum of 2, 4, 6, 8, and 10.

```vb
Sub TwosTotal() 
    For j = 2 To 10 Step 2 
        total = total + j 
    Next j 
    MsgBox "The total is " & total 
End Sub
```

To decrease the counter variable, use a negative **Step** value. To decrease the counter variable, you must specify an end value that is less than the start value. In the following example, the counter variable `myNum` is decreased by 2 each time the loop repeats. When the loop is finished, `total` is the sum of 16, 14, 12, 10, 8, 6, 4, and 2.

```vb
Sub NewTotal() 
    For myNum = 16 To 2 Step -2 
        total = total + myNum 
    Next myNum 
    MsgBox "The total is " & total 
End Sub
```

> [!NOTE] 
> It's not necessary to include the counter variable name after the **Next** statement. In the preceding examples, the counter variable name was included for readability.

You can exit a **For...Next** statement before the counter reaches its end value by using the **[Exit For](../../reference/user-interface-help/exit-statement.md)** statement. For example, when an error occurs, use the **Exit For** statement in the **True** statement block of either an **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statement or a **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement that specifically checks for the error. If the error doesn't occur, the **If…Then…Else** statement is **False**, and the loop will continue to run as expected.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
