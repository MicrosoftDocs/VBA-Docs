---
title: AutoRecover object (Excel)
keywords: vbaxl10.chm695072
f1_keywords:
- vbaxl10.chm695072
ms.prod: excel
api_name:
- Excel.AutoRecover
ms.assetid: 02fb24e7-4823-7e52-79d7-3d2726f31227
ms.date: 03/29/2019
localization_priority: Normal
---


# AutoRecover object (Excel)

Represents the automatic recovery features of a workbook. 


## Remarks

Properties for the **AutoRecover** object determine the path and time interval for backing up all files.

Use the **[AutoRecover](Excel.Application.AutoRecover.md)** property of the **Application** object to return an **AutoRecover** object.

Use the **Path** property of the **AutoRecover** object to set the path for where the AutoRecover file will be saved.


## Example

The following example sets the path of the AutoRecover file to drive C.

```vb
Sub SetPath() 
 
 Application.AutoRecover.Path = "C:\" 
 
End Sub
```

<br/>

Use the **Time** property of the **AutoRecover** object to set the time interval for backing up all files. Units for the **Time** property are in minutes.

```vb
Sub SetTime() 
 
 Application.AutoRecover.Time = 5 
 
End Sub
```


## Properties

- [Application](Excel.AutoRecover.Application.md)
- [Creator](Excel.AutoRecover.Creator.md)
- [Enabled](Excel.AutoRecover.Enabled.md)
- [Parent](Excel.AutoRecover.Parent.md)
- [Path](Excel.AutoRecover.Path.md)
- [Time](Excel.AutoRecover.Time.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]