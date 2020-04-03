---
title: Task.SendWindowMessage method (Word)
keywords: vbawd10.chm159514638
f1_keywords:
- vbawd10.chm159514638
ms.prod: word
api_name:
- Word.Task.SendWindowMessage
ms.assetid: 3c4793b4-30cd-e27e-2b9f-cc5187304ddc
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.SendWindowMessage method (Word)

Sends a Windows message and its associated parameters to the specified task.


## Syntax

_expression_. `SendWindowMessage`( `_Message_` , `_wParam_` , `_IParam_` )

_expression_ Required. A variable that represents a '[Task](Word.Task.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Message_|Required| **Long**|A hexadecimal number that corresponds to the message you want to send. If you have the Microsoft Platform Software Development Kit, you can look up the name of the message in the header files (Winuser.h, for example) to find the associated hexadecimal number (precede the hexadecimal value with &h).|
| _wParam_|Required| **Long**|Parameters appropriate for the message you are sending. For information about what these values represent, see the reference topic for that message in the documentation included with the Microsoft Platform Software Development Kit, available on MSDN. To retrieve the appropriate values, you may need to use the Spy tool (which comes with the kit).|

## Example

If Notepad is running, this example displays the **About** dialog box (in Notepad) by sending a WM_COMMAND message to Notepad. The **SendWindowMessage** method is used to send the WM_COMMAND message (111 is the hexadecimal value for WM_COMMAND), with the parameters 11 and 0. The Spy tool was used to determine the **wParam** and **lParam** values.


```vb
Dim taskLoop As Task 
 
For Each taskLoop In Tasks 
 If InStr(taskLoop.Name, "Notepad") > 0 Then 
 taskLoop.Activate 
 taskLoop.SendWindowMessage &h111, 11, 0 
 End If 
Next taskLoop
```


## See also


[Task Object](Word.Task.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]