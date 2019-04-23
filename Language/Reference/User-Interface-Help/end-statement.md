---
title: End statement (VBA)
keywords: vblr6.chm1008904
f1_keywords:
- vblr6.chm1008904
ms.prod: office
ms.assetid: 5cbb1c20-2afa-782e-52bb-7aafc604a927
ms.date: 12/03/2018
localization_priority: Normal
---


# End statement

Ends a [procedure](../../Glossary/vbe-glossary.md#procedure) or block.

## Syntax

**End** <br/>
**End Function** <br/>
**End If** <br/>
**End Property** <br/>
**End Select** <br/>
**End Sub** <br/>
**End Type** <br/>
**End With** 

<br/>

The **End** statement syntax has these forms:

|Statement|Description|
|:-----|:-----|
|**End**|Terminates execution immediately. Never required by itself but may be placed anywhere in a procedure to end code execution, close files opened with the **[Open](open-statement.md)** statement, and to clear [variables](../../Glossary/vbe-glossary.md#variable).|
|**End Function**|Required to end a **[Function](function-statement.md)** statement.|
|**End If**|Required to end a block **[If…Then…Else](ifthenelse-statement.md)** statement.|
|**End Property**|Required to end a **[Property Let](property-let-statement.md)**, **[Property Get](property-get-statement.md)**, or **[Property Set](property-set-statement.md)** procedure.|
|**End Select**|Required to end a **[Select Case](select-case-statement.md)** statement.|
|**End Sub**|Required to end a **[Sub](sub-statement.md)** statement.|
|**End Type**|Required to end a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) definition (**[Type](type-statement.md)** statement).|
|**End With**|Required to end a **[With](with-statement.md)** statement.|

## Remarks 

When executed, the **End** statement resets all [module-level](../../Glossary/vbe-glossary.md#module-level) variables and all static local variables in all [modules](../../Glossary/vbe-glossary.md#module). To preserve the value of these variables, use the **[Stop](stop-statement.md)** statement instead. You can then resume execution while preserving the value of those variables.

> [!NOTE] 
> The **End** statement stops code execution abruptly, without invoking the Unload, QueryUnload, or Terminate event, or any other Visual Basic code. Code you have placed in the Unload, QueryUnload, and Terminate events of [forms](../../Glossary/vbe-glossary.md#form) and [class modules](../../Glossary/vbe-glossary.md#class-module) is not executed. Objects created from class modules are destroyed, files opened by using the **Open** statement are closed, and memory used by your program is freed. Object references held by other programs are invalidated.

The **End** statement provides a way to force your program to halt. For normal termination of a Visual Basic program, you should unload all forms. Your program closes as soon as there are no other programs holding references to objects created from your public class modules and no code executing.

## Example

This example uses the **End** statement to end code execution if the user enters an invalid password.


```vb
Sub Form_Load 
  Dim Password, Pword 
  PassWord = "Swordfish" 
  Pword = InputBox("Type in your password") 
  If Pword <> PassWord Then 
    MsgBox "Sorry, incorrect password" 
    End
  End If
End Sub
```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
