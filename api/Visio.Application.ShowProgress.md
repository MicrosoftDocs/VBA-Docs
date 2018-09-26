---
title: Application.ShowProgress Property (Visio)
keywords: vis_sdr.chm10014370
f1_keywords:
- vis_sdr.chm10014370
ms.prod: visio
api_name:
- Visio.Application.ShowProgress
ms.assetid: 4dcfcec7-d652-0b52-a4e8-a43122e72988
ms.date: 06/08/2017
---


# Application.ShowProgress Property (Visio)

Determines whether a progress indicator is shown while performing certain operations. Read/write.


## Syntax

 _expression_. `ShowProgress`

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Return value

Integer


## Remarks

If you want to perform an operation, such as printing, that typically displays a progress indicator but you don't want the progress indicator to appear, set the  **ShowProgress** property to **False** (0). By default, the **ShowProgress** property is **True** (non-zero).

In most cases you should restore the setting to its prior value when you've completed the operation.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ShowProgress** property of the **Application** object. It switches the display of the progress indicator on and off and displays the state of the property in the Immediate window.


```vb
 
Sub ShowProgress_Example() 
 
 'Create a variable to hold the current state 
 'of the progress indicator. 
 Dim intProgress As Integer 
 
 'Display the current state of the progress indicator. 
 Debug.Print "Initial state: " & Application.ShowProgress 
 
 'Get the state of the progress indicator. 
 intProgress = Application.ShowProgress 
 
 'Turn off the progress indicator. 
 Application.ShowProgress = False 
 
 'Display the status again. 
 Debug.Print "Changed state: " & Application.ShowProgress 
 
 'Now restore the progress indicator to its original state. 
 Application.ShowProgress = intProgress 
 
 'Display the status again. 
 Debug.Print "Restored state: " & Application.ShowProgress 
 
End Sub
```


