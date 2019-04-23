---
title: CommandBars.DisableCustomize property (Office)
keywords: vbaof11.chm2016
f1_keywords:
- vbaof11.chm2016
ms.prod: office
api_name:
- Office.CommandBars.DisableCustomize
ms.assetid: cbebdaa7-2e8d-af73-fd18-03b3b11f98ac
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.DisableCustomize property (Office)

Is **True** if toolbar customization is disabled. Read/write.


## Syntax

_expression_.**DisableCustomize**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Example

The following example switches the **DisableCustomize** property on or off.


```vb
Sub ToggleCustomize() 
 With Application.CommandBars 
 If .DisableCustomize = True Then 
 .DisableCustomize = False 
 Else 
 .DisableCustomize = True 
 End If 
 End With 
End Sub
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]