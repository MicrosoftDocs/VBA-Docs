---
title: Application.BoxCellEdit method (Project)
keywords: vbapj.chm2393
f1_keywords:
- vbapj.chm2393
ms.prod: project-server
api_name:
- Project.Application.BoxCellEdit
ms.assetid: 27063852-3dc4-57b2-c82a-6210674810ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BoxCellEdit method (Project)

Sets the properties of an individual cell in a data template for a Network Diagram view (PERT chart).


## Syntax

_expression_. `BoxCellEdit`( `_Name_`, `_Cell_`, `_FieldName_`, `_Font_`, `_FontSize_`, `_FontColor_`, `_Bold_`, `_Italic_`, `_Underline_`, `_HorizontalAlignment_`, `_VerticalAlignment_`, `_TextLineLimit_`, `_ShowLabel_`, `_Label_`, `_DateFormat_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of the data template containing the cell to edit.|
| _Cell_|Required|**Long**|The cell to edit. Can be one of the  **[PjCell](Project.PjCell.md)** constants.|
| _FieldName_|Optional|**Long**|The name of the field to display in the cell. Can be one of the  **[PjField](Project.PjField.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _FontSize_|Optional|**Integer**|The size of the font, in points.|
| _FontColor_|Optional|**Long**|The color of the font. Can be one of the  **[PjColor](Project.PjColor.md)** constants.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _HorizontalAlignment_|Optional|**Long**|Specifies the horizontal alignment of text in the cell. Can be one of the  **[PjAlignment](Project.PjAlignment.md)** constants.|
| _VerticalAlignment_|Optional|**Long**|Specifies the horizontal alignment of text in the cell. Can be one of the  **[PjVerticalAlignment](Project.PjVerticalAlignment.md)** constants.|
| _TextLineLimit_|Optional|**Long**|Specifies the limit for the number of lines of text in the cell. Values can be 1, 2, or 3. |
| _ShowLabel_|Optional|**Boolean**|**True** if a label is shown in the cell; otherwise, **False**.|
| _Label_|Optional|**String**|Specifies the cell label.|
| _DateFormat_|Optional|**Long**|Specifies the date format for the cell when  **FieldName** is a date field. Can be one of the **[PjDateFormat](Project.PjDateFormat.md)** constants.|

## Return value

 **Boolean**


## Remarks

Using the  **BoxCellEdit** method with only the _Name_ and _Cell_ arguments has no effect.

To edit a data template cell where the text color can be an RGB value, use the [BoxCellEditEx](Project.Application.BoxCellEditEx.md) method.


## Example

The following example modifies a copy of the  **Critical** data template named **Test Critical**. The macro removes the fourth row of cells and sets the fourth cell in the third row to show the  **Actual Cost** field and label in a green color.


```vb
Sub ModifyCriticalDataTemplate() 
    Application.BoxCellLayout Name:="Test Critical", CellRows:=3, MergeCells:=True 
 
    Application.BoxCellEdit Name:="Test Critical", Cell:=pjCell4_3, _ 
        FieldName:=PjField.pjTaskActualCost, Font:="Arial", FontSize:="8", FontColor:=PjColor.pjGreen, _ 
        Bold:=False, Italic:=False, Underline:=False, HorizontalAlignment:=pjLeft, _ 
        VerticalAlignment:=pjMiddle, TextLineLimit:=1, ShowLabel:=True, Label:="Cost" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]