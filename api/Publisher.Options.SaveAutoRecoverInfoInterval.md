---
title: Options.SaveAutoRecoverInfoInterval property (Publisher)
keywords: vbapb10.chm1048600
f1_keywords:
- vbapb10.chm1048600
ms.prod: publisher
api_name:
- Publisher.Options.SaveAutoRecoverInfoInterval
ms.assetid: 3d6a6c4f-7e2b-18ff-67a4-20dee4fbcf5b
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.SaveAutoRecoverInfoInterval property (Publisher)

Returns or sets a **Long** that represents the time interval in minutes for automatically saving a publication for recovery if the application is unexpectedly shut down. Read/write.


## Syntax

_expression_.**SaveAutoRecoverInfoInterval**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Long


## Example

This example enables the global auto recovery option and sets the save interval to every five minutes.

```vb
Sub SetAutoRecoverInfo() 
 With Options 
 .SaveAutoRecoverInfo = True 
 .SaveAutoRecoverInfoInterval = 5 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]