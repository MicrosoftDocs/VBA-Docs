---
title: AutoRecover Object (Excel)
keywords: vbaxl10.chm695072
f1_keywords:
- vbaxl10.chm695072
ms.prod: excel
api_name:
- Excel.AutoRecover
ms.assetid: 02fb24e7-4823-7e52-79d7-3d2726f31227
ms.date: 06/08/2017
---


# AutoRecover Object (Excel)

Represents the automatic recovery features of a workbook. 


## Remarks

Properties for the  **AutoRecover** object determine the path and time interval for backing up all files.

Use the  **[AutoRecover](Excel.Application.AutoRecover.md)** property of the **[Application](Excel.Application(object).md)** object to return an **AutoRecover** object.

Use the  **[Path](Excel.AutoRecover.Path.md)** property of the **AutoRecover** object to set the path for where the AutoRecover file will be saved.


## Example

The following example sets the path of the AutoRecover file to drive C.


```vb
Sub SetPath() 
 
 Application.AutoRecover.Path = "C:\" 
 
End Sub
```

Use the  **[Time](Excel.AutoRecover.Time.md)** property of the **AutoRecover** object to set the time interval for backing up all files.


 **Note**  Units for the  **Time** property are in minutes.




```vb
Sub SetTime() 
 
 Application.AutoRecover.Time = 5 
 
End Sub
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

