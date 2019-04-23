---
title: Using For Each...Next statements (VBA)
keywords: vbcn6.chm1076683
f1_keywords:
- vbcn6.chm1076683
ms.prod: office
ms.assetid: 76df8944-219a-c28b-c449-39a3108c11be
ms.date: 12/26/2018
localization_priority: Normal
---


# Using For Each...Next statements

**[For Each...Next](../../reference/user-interface-help/for-eachnext-statement.md)** statements repeat a block of [statements](../../Glossary/vbe-glossary.md#statement) for each [object](../../Glossary/vbe-glossary.md#object) in a [collection](../../Glossary/vbe-glossary.md#collection) or each element in an [array](../../Glossary/vbe-glossary.md#array). Visual Basic automatically sets a [variable](../../Glossary/vbe-glossary.md#variable) each time the loop runs. For example, the following [procedure](../../Glossary/vbe-glossary.md#procedure) closes all forms except the form containing the procedure that's running.

```vb
Sub CloseForms() 
 For Each frm In Application.Forms 
 If frm.Caption <> Screen. ActiveForm.Caption Then frm.Close 
 Next 
End Sub
```

The following code loops through each element in an array and sets the value of each to the value of the index variable I.

```vb
Dim TestArray(10) As Integer, I As Variant 
For Each I In TestArray 
 TestArray(I) = I 
Next I 

```


## Looping through a range of cells

Use a **For Each...Next** loop to loop through the cells in a range. The following procedure loops through the range A1:D10 on Sheet1 and sets any number whose absolute value is less than 0.01 to 0 (zero).

```vb
Sub RoundToZero() 
 For Each myObject in myCollection 
 If Abs(myObject.Value) < 0.01 Then myObject.Value = 0 
 Next 
End Sub
```

## Exiting a For Each...Next loop before it is finished

You can exit a **For Each...Next** loop by using the **[Exit For](../../reference/user-interface-help/exit-statement.md)** statement. For example, when an error occurs, use the **Exit For** statement in the **True** statement block of either an **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statement or a **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement that specifically checks for the error. If the error does not occur, the **If…Then…Else** statement is **False** and the loop continues to run as expected.

The following example tests for the first cell in the range A1:B5 that does not contain a number. If such a cell is found, a message is displayed and **Exit For** exits the loop.

```vb
Sub TestForNumbers() 
 For Each myObject In MyCollection 
 If IsNumeric(myObject.Value) = False Then 
 MsgBox "Object contains a non-numeric value." 
 Exit For 
 End If 
 Next c 
End Sub
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
