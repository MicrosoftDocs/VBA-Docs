---
title: Form.CommandExecute event (Access)
keywords: vbaac10.chm13676
f1_keywords:
- vbaac10.chm13676
ms.prod: access
api_name:
- Access.Form.CommandExecute
ms.assetid: b4b3bc8e-3e95-5120-ed7e-e17b2f8f23ba
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandExecute event (Access)

Occurs after the specified command is executed. Use this event when you want to execute a set of commands after a particular command is executed.


## Syntax

_expression_.**CommandExecute** (_Command_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**Variant**|The command that is executed.|


## Return value

Nothing


## Remarks

The **OCCommandId**, **ChartCommandIdEnum**, and **PivotCommandId** constants contain lists of the supported commands for each of the Microsoft Office web components.


## Example

The following example demonstrates the syntax for a subroutine that traps the **CommandExecute** event.

```vb
Private Sub Form_CommandExecute(ByVal Command As Variant) MsgBox "The command specified by " _ 
 & Command.Name & " has been executed." 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]