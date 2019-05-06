---
title: Writing executable statements (VBA)
keywords: vbcn6.chm1076694
f1_keywords:
- vbcn6.chm1076694
ms.prod: office
ms.assetid: 822a0e4e-687d-9f38-7b70-352f3ee10da1
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing executable statements

An executable [statement](../../Glossary/vbe-glossary.md#statement) initiates action. It can execute a [method](../../Glossary/vbe-glossary.md#method) or function, and it can loop or branch through blocks of code. Executable statements often contain mathematical or conditional operators.

The following example uses a **For Each...Next** statement to iterate through each cell in a range named _MyRange_ on Sheet1 of an active Microsoft Excel workbook. The variable `c` is a cell in the collection of cells contained in _MyRange_.

```vb
Sub ApplyFormat() 
Const limit As Integer = 33 
For Each c In Worksheets("Sheet1").Range("MyRange").Cells 
    If c.Value > limit Then 
        With c.Font 
            .Bold = True 
            .Italic = True 
        End With 
    End If 
Next c 
MsgBox "All done!" 
End Sub
```

The **If...Then...Else** statement in the example checks the value of the cell. If the value is greater than 33, the **With** statement sets the **Bold** and **Italic** properties of the **Font** object for that cell. **If...Then...Else** statements end with **End If**. The **With** statement can save typing because the statements it contains are automatically executed on the object following the **With** keyword.

The **Next** statement calls the next cell in the collection of cells contained in _MyRange_.

The **MsgBox** function (which displays a built-in Visual Basic dialog box) displays a message indicating that the **Sub** procedure has finished running.

## See also

- [Statements](../../reference/statements.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]