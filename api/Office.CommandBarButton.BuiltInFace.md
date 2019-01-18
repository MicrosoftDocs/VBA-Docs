---
title: CommandBarButton.BuiltInFace property (Office)
keywords: vbaof11.chm6001
f1_keywords:
- vbaof11.chm6001
ms.prod: office
api_name:
- Office.CommandBarButton.BuiltInFace
ms.assetid: 47c82878-17ea-b6ff-e841-c9f07342c8a3
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.BuiltInFace property (Office)

Is **True** if the face of a command bar button control is its original built-in face. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BuiltInFace**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

This property can only be set to **True**, which will reset the face to the built-in face. Read/write **Boolean**.


## Example

This example determines whether the face of the first control on the command bar named **Custom** is its built-in button face. If it is, the example copies the button face to the Clipboard.


```vb
Set myControl = CommandBars("My Custom Bar").Controls(1) 
With myControl 
    If .BuiltInFace = True Then .CopyFace 
End With
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]