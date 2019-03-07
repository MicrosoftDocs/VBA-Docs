---
title: Form.CommandEnabled event (Access)
keywords: vbaac10.chm13675
f1_keywords:
- vbaac10.chm13675
ms.prod: access
api_name:
- Access.Form.CommandEnabled
ms.assetid: 4a9ff0dc-5ed2-e841-97d3-a1c4a7ed4d42
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandEnabled event (Access)

Occurs when the specified Microsoft Office web component determines whether the specified command is enabled.


## Syntax

_expression_.**CommandEnabled** (_Command_, _Enabled_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**Variant**| The command that has been verified as being enabled.|
| _Enabled_|Required|**Object**|Set the **Value** property of this object to **False** to disable the command.|

## Return value

Nothing


## Remarks

The **OCCommandId**, **ChartCommandIdEnum**, and **PivotCommandId** constants contain lists of the supported commands for each of the Microsoft Office web components.


## Example

The following example demonstrates the syntax for a subroutine that traps the **CommandEnabled** event.

```vb
Private Sub Form_CommandEnabled( _ 
 ByVal Command As Variant, ByVal Enabled As Object) Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Disable the command?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Enabled.Value = False 
 Else 
 Enabled.Value = True 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]