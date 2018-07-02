---
title: CommandBars.DisplayFonts Property (Office)
keywords: vbaof11.chm2015
f1_keywords:
- vbaof11.chm2015
ms.prod: office
api_name:
- Office.CommandBars.DisplayFonts
ms.assetid: 25a9ede7-3575-6706-406d-a5b656cd965e
ms.date: 06/08/2017
---


# CommandBars.DisplayFonts Property (Office)

Is  **True** if the font names in the **Font** box are displayed in their actual fonts. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `DisplayFonts`

 _expression_ A variable that represents a [CommandBars](./Office.CommandBars.md) object.


## Example

This example sets three options for all command bars in Microsoft Office, including custom command bars and the controls on them.


```vb
With CommandBars 
    .LargeButtons = True  
    .DisplayFonts = True  
    .AdaptiveMenus = True  
End With
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/commandbars-members-office.md)

