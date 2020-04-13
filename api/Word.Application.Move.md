---
title: Application.Move method (Word)
keywords: vbawd10.chm158335336
f1_keywords:
- vbawd10.chm158335336
ms.prod: word
api_name:
- Word.Application.Move
ms.assetid: 030b6ae1-50bd-8d3e-e760-509c54a6e152
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Move method (Word)

Positions a task window or the active document window.


## Syntax

_expression_. `Move`( `_Left_` , `_Top_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Long**|The horizontal screen position of the specified window.|
| _Top_|Required| **Long**|The vertical screen position of the specified window.|

## Example

This example starts the Calculator application (Calc.exe) and uses the **Move** method to reposition the application window.


```vb
Shell "Calc.exe" 
With Tasks("Calculator") 
 .WindowState = wdWindowStateNormal 
 .Move Top:=50, Left:=50 
End With
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]