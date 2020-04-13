---
title: Application.FilePageSetupCalendarText method (Project)
keywords: vbapj.chm2371
f1_keywords:
- vbapj.chm2371
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupCalendarText
ms.assetid: 279e4f0e-f2fb-0822-bf75-700b365c301d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FilePageSetupCalendarText method (Project)

Formats the text of calendar views for printing.


## Syntax

_expression_. `FilePageSetupCalendarText`( `_Name_`, `_Item_`, `_Font_`, `_Size_`, `_Bold_`, `_Italic_`, `_Underline_`, `_Color_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the calendar to edit.|
| _Item_|Optional|**Long**|The text item to format. Can be one of the **[PjPageSetupCalendarItem](Project.PjPageSetupCalendarItem.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the text. Can be one of the **[PjColor](Project.PjColor.md)** constants.|

## Return value

 **Boolean**


## Remarks

Using the **FilePageSetupCalendarText** method without any arguments displays the **Text Styles** dialog box.


> [!NOTE] 
>  **FilePageSetupCalendarText** works only for printing calendar views.

To format calendar text where  _Color_ can be a hexadecimal RGB value, use the **[FilePageSetupCalendarTextEx](Project.Application.FilePageSetupCalendarTextEx.md)** method.


## Example

The following example formats monthly titles in red for printing.


```vb
Sub File_PageSetupCalendarText() 
 
 'Activate the Calendar view. 
 ViewApply Name:="&Calendar" 
 FilePageSetupCalendarText Item:=pjMonthlyTitles, Color:=pjRed 
 FilePrint 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]