---
title: System object (Word)
keywords: vbawd10.chm2357
f1_keywords:
- vbawd10.chm2357
ms.prod: word
api_name:
- Word.System
ms.assetid: db15d780-3bbc-9515-a988-ea798777496f
ms.date: 06/08/2017
localization_priority: Normal
---


# System object (Word)

Contains information about the computer system.


## Remarks

Use the **System** property to return the **System** object. If the operating system is Windows, the following example makes a network connection to \\Project\Info.


```vb
If System.OperatingSystem = "Windows" Then 
 System.Connect Path:="\\Project\Info" 
End If
```

The following example displays the current screen resolution (for example, "1024 x 768").




```vb
horz = System.HorizontalResolution 
vert = System.VerticalResolution 
MsgBox "Resolution = " & horz & " x " & vert
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]