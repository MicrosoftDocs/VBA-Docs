---
title: Selection.MoveRight method (Word)
keywords: vbawd10.chm158663157
f1_keywords:
- vbawd10.chm158663157
ms.prod: word
api_name:
- Word.Selection.MoveRight
ms.assetid: fcac96c7-7189-87b2-d800-9d161edb1e09
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.MoveRight method (Word)

Moves the selection to the right and returns the number of units it has been moved.


## Syntax

_expression_. `MoveRight`( `_Unit_` , `_Count_` , `_Extend_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which the selection is to be moved.The default value is  **wdCharacter**.|
| _Count_|Optional| **Variant**|The number of units the selection is to be moved. The default value is 1.|
| _Extend_|Optional| **Variant**|Can be either  **wdMove** or **wdExtend**. If **wdMove** is used, the selection is collapsed to the endpoint and moved to the right. If **wdExtend** is used, the selection is extended to the right. The default value is **wdMove**.|

## Return value

Long


## Remarks

When the Unit is  **wdCell**, the Extend argument can only be **wdMove**.


## Example

This example moves the selection before the previous field and then selects the field.


```vb
With Selection 
 Set MyRange = .GoTo(wdGoToField, wdGoToPrevious) 
 .MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend 
 If Selection.Fields.Count = 1 Then Selection.Fields(1).Update 
End With
```

This example moves the selection one character to the right. If the move is successful, MoveRight returns 1.




```vb
If Selection.MoveRight = 1 Then MsgBox "Move was successful"
```

This example moves the selection to the next table cell.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.MoveRight Unit:=wdCell, Count:=1, Extend:=wdMove 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
