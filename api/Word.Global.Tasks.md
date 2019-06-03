---
title: Global.Tasks property (Word)
keywords: vbawd10.chm163119132
f1_keywords:
- vbawd10.chm163119132
ms.prod: word
api_name:
- Word.Global.Tasks
ms.assetid: e6a89660-adfd-a8f0-6322-ac232ba3dce2
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Tasks property (Word)

Returns a  **Tasks** collection that represents all the applications that are running.


## Syntax

_expression_. `Tasks`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the calculator. If the calculator is not already running, then Word starts the task and then displays the calculator.


```vb
If Tasks.Exists("Calculator") Then 
 With Tasks("Calculator") 
 .Activate 
 .WindowState = wdWindowStateNormal 
 End With 
Else 
 Shell "calc.exe" 
 Tasks("Calculator").WindowState = wdWindowStateNormal 
End If
```

This example checks to see whether Microsoft Excel is currently running. If the task is running, the example activates Microsoft Excel; otherwise, a message box is displayed.




```vb
If Tasks.Exists("Microsoft Excel") = True Then 
 With Tasks("Microsoft Excel") 
 .Activate 
 .WindowState = wdWindowStateMaximize 
 End With 
Else 
 Msgbox "Microsoft Excel is not currently running." 
End If
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]