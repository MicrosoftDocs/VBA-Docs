---
title: Form.CommandChecked event (Access)
keywords: vbaac10.chm13674
f1_keywords:
- vbaac10.chm13674
ms.prod: access
api_name:
- Access.Form.CommandChecked
ms.assetid: ec30f538-bbd2-9935-1ad9-5210f457b15f
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandChecked event (Access)

Occurs when the specified Microsoft Office web component determines whether the specified command is selected.


## Syntax

_expression_.**CommandChecked** (_Command_, _Checked_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**Variant**| The command that has been verified as being selected.|
| _Checked_|Required|**Object**| Set the **Value** property of this object to **False** to clear the command.|

## Return value

Nothing


## Remarks

The **OCCommandId**, **ChartCommandIdEnum**, and **PivotCommandId** constants contain lists of the supported commands for each of the Microsoft Office web components.


## Example

The following example demonstrates the syntax for a subroutine that traps the **CommandChecked** event.

```vb
Private Sub Form_CommandChecked( _ 
 ByVal Command As Variant, ByVal Checked As Object) Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Uncheck the command?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Checked.Value = False 
 Else 
 Checked.Value = True 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]