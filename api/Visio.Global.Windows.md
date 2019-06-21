---
title: Global.Windows property (Visio)
keywords: vis_sdr.chm12814665
f1_keywords:
- vis_sdr.chm12814665
ms.prod: visio
api_name:
- Visio.Global.Windows
ms.assetid: d86b6db0-702c-9058-03a7-b457388ebfd3
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Windows property (Visio)

Returns the  **Windows** collection for a Microsoft Visio instance or window. Read-only.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents a **[Global](Visio.Global.md)** object.


## Return value

Windows


## Example

This Microsoft Visual Basic macro gets the  **Windows** collection of the **Application** object and prints the ID of each window in the collection in the Immediate window.


```vb
Public Sub Windows_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoWindows As Visio.Windows 
 Dim intCounter As Integer 
 
 'Get the Windows collection. 
 Set vsoApplication = Application 
 Set vsoWindows = vsoApplication.Windows 
 
 For intCounter = 1 To vsoWindows.Count 
 Debug.Print vsoWindows.Item(intCounter).ID 
 Next intCounter 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]