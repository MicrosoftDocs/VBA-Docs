---
title: Speech.Direction property (Excel)
keywords: vbaxl10.chm718074
f1_keywords:
- vbaxl10.chm718074
ms.prod: excel
api_name:
- Excel.Speech.Direction
ms.assetid: 8cbedcb3-2d91-b9c1-c1ae-6f06cd8d442b
ms.date: 06/08/2017
localization_priority: Normal
---


# Speech.Direction property (Excel)

Returns or sets the order in which the cells will be spoken. The value of the  **Direction** property is an **[xlSpeakDirection](Excel.XlSpeakDirection.md)** constant. Read/write.


## Syntax

_expression_. `Direction`

_expression_ A variable that represents a [Speech](./Excel.Speech.md) object.


## Remarks





| **xlSpeakDirection** can be one of these **xlSpeakDirection** constants.|
| **xlSpeakByColumns**|
| **xlSpeakByRows**|

## Example

In this example, Microsoft Excel determines the speech direction and notifies the user.


```vb
Sub CheckSpeechDirection() 
 
 ' Notify user of speech direction. 
 If Application.Speech.Direction = xlSpeakByColumns Then 
 MsgBox "The speech direction is set to speak by columns." 
 Else 
 MsgBox "The speech direction is set to speak by rows." 
 End If 
 
End Sub
```


## See also


[Databar Object](Excel.Databar.md)
[Speech Object](Excel.Speech.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]