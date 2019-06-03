---
title: CommandBars.FindControls method (Office)
keywords: vbaof11.chm2014
f1_keywords:
- vbaof11.chm2014
ms.prod: office
api_name:
- Office.CommandBars.FindControls
ms.assetid: 79c46884-816d-def6-2bff-85b59b0831ea
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.FindControls method (Office)

Gets the **CommandBarControls** collection that fits the specified criteria.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**FindControls** (_Type_, _Id_, _Tag_, _Visible_)

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|Is one of the **[MsoControlType](office.msocontroltype.md)** constants specifying the type of control.|
| _Id_|Optional|**Variant**|The control's identifier.|
| _Tag_|Optional|**Variant**|The control's tag value.|
| _Visible_|Optional|**Variant**|**True** to include only visible command bar controls in the search. The default value is **False**.|

## Return value

CommandBarControls


## Remarks

If no controls that fit the criteria are found, the **FindControls** method returns **Nothing**.


## Example

This example uses the **FindControls** method to return all members of the **CommandBars** collection that have an ID of 18 and displays (in a message box) the number of controls that meet the search criteria.


```vb
Dim myControls As CommandBarControls 
Set myControls = CommandBars.FindControls(Type:=msoControlButton, ID:=18) 
MsgBox "There are " & myControls.Count & _ 
    " controls that meet the search criteria."
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]