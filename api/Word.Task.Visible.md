---
title: Task.Visible property (Word)
keywords: vbawd10.chm159514630
f1_keywords:
- vbawd10.chm159514630
ms.prod: word
api_name:
- Word.Task.Visible
ms.assetid: cc1bb50d-c49d-9230-83ad-940c53c89220
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Visible property (Word)

 **True** if the specified object is visible. Read/write **Boolean**.


## Syntax

_expression_.**Visible**

_expression_ Required. A variable that represents a '[Task](Word.Task.md)' object.


## Remarks

For any object, some methods and properties may be unavailable if the **Visible** property is **False**.


## Example

This example hides the Calculator, if it is running. If it is not running, a message is displayed.


```vb
If Tasks.Exists("Calculator") Then 
 Tasks("Calculator").Visible = False 
Else 
 Msgbox "Calculator is not running." 
End If
```


## See also


[Task Object](Word.Task.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]