---
title: Options.SaveAutoRecoverInfo property (Publisher)
keywords: vbapb10.chm1048599
f1_keywords:
- vbapb10.chm1048599
ms.prod: publisher
api_name:
- Publisher.Options.SaveAutoRecoverInfo
ms.assetid: 1cbb7960-8995-37f4-5989-01b97152269f
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.SaveAutoRecoverInfo property (Publisher)

**True** if Microsoft Publisher automatically saves publications for recovery if the application is unexpectedly shut down. Read/write **Boolean**.


## Syntax

_expression_.**SaveAutoRecoverInfo**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Remarks

Use the **[SaveAutoRecoverInfoInterval](Publisher.Options.SaveAutoRecoverInfoInterval.md)** property to specify how often auto recovery saves occur.


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