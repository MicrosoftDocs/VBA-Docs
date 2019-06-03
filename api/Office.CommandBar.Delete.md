---
title: CommandBar.Delete method (Office)
keywords: vbaof11.chm3004
f1_keywords:
- vbaof11.chm3004
ms.prod: office
api_name:
- Office.CommandBar.Delete
ms.assetid: 6976f273-dbd4-5f3d-52ef-0d6d5cc886c9
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Delete method (Office)

Deletes the **CommandBar** object from the collection.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Delete**

_expression_ Required. A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

For the **Scripts** collection, using the **Delete** method removes all scripts from the specified Microsoft Word document, Excel worksheet, or PowerPoint slide. A script anchor is represented by a shape in the host application. Therefore, the **Shape** object associated with each script anchor of type **msoScriptAnchor** is deleted from the **Shapes** collection in Excel and PowerPoint and from the **InlineShapes** and **Shapes** collections in Word.


## Example

This example deletes all custom command bars that are not visible.


```vb
foundFlag = False  
delBars = 0 
For Each bar In CommandBars 
    If (bar.BuiltIn = False) And _ 
    (bar.Visible = False) Then 
        bar.Delete 
        foundFlag =   
        delBars = delBars + 1 
    End If 
Next bar 
If Not foundFlag Then 
    MsgBox "No command bars have been deleted." 
Else 
    MsgBox delBars & " custom bar(s) deleted." 
End If
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]