---
title: Application.TransitionNavigKeys Property (Excel)
keywords: vbaxl10.chm133220
f1_keywords:
- vbaxl10.chm133220
ms.prod: excel
api_name:
- Excel.Application.TransitionNavigKeys
ms.assetid: 261afa51-44f7-4527-9145-b542cc68d812
ms.date: 06/08/2017
---


# Application.TransitionNavigKeys Property (Excel)

 **True** if transition navigation keys are active. Read/write **Boolean** .


## Syntax

 _expression_. `TransitionNavigKeys`

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


## Example

This example displays the current state of the  **Transition navigation keys** option.


```vb
If Application.TransitionNavigKeys Then 
 keyState = "On" 
Else 
 keyState = "Off" 
End If 
MsgBox "The Transition Navigation Keys option is " & keyState
```


## See also


[Application Object](Excel.Application(object).md)

