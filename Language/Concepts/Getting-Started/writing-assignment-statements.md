---
title: Writing assignment statements (VBA)
keywords: vbcn6.chm1076692
f1_keywords:
- vbcn6.chm1076692
ms.prod: office
ms.assetid: 7699bec2-c5a2-6f35-3ec0-8aa7cefa622d
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing assignment statements

Assignment statements assign a value or [expression](../../Glossary/vbe-glossary.md#expression) to a [variable](../../Glossary/vbe-glossary.md#variable) or [constant](../../Glossary/vbe-glossary.md#constant). Assignment statements always include an equal sign (**=**). 

The following example assigns the return value of the **InputBox** function to the variable.

```vb
Sub Question() 
 Dim yourName As String 
 yourName = InputBox("What is your name?") 
 MsgBox "Your name is " & yourName 
End Sub
```


The **Let** statement is optional and is usually omitted. For example, the preceding assignment statement can be written.

```vb
Let yourName = InputBox("What is your name?"). 

```

The **[Set](../../reference/user-interface-help/set-statement.md)** statement is used to assign an object to a variable that has been declared as an object. The **Set** keyword is required. In the following example, the **Set** statement assigns a range on Sheet1 to the object variable `myCell`.

```vb
Sub ApplyFormat() 
Dim myCell As Range 
Set myCell = Worksheets("Sheet1").Range("A1") 
 With myCell.Font 
 .Bold = True 
 .Italic = True 
 End With 
End Sub
```

Statements that set [property](../../Glossary/vbe-glossary.md#property) values are also assignment statements. The following example sets the **Bold** property of the **Font** object for the active cell.

```vb
ActiveCell.Font.Bold = True 

```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]