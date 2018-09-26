---
title: Application.BoxLayoutEx Method (Project)
keywords: vbapj.chm2157
f1_keywords:
- vbapj.chm2157
ms.prod: project-server
api_name:
- Project.Application.BoxLayoutEx
ms.assetid: 40c80e1c-6763-172d-c48a-0ec7c1fa2412
ms.date: 06/08/2017
---


# Application.BoxLayoutEx Method (Project)

Specifies the layout of boxes in the active Network Diagram view (PERT chart), where the background color can be specified as a hexadecimal value.


## Syntax

 _expression_. `BoxLayoutEx`( ` _LayoutMode_`, ` _LayoutScheme_`, ` _SummaryPrecedence_`, ` _RowAlignment_`, ` _ColumnAlignment_`, ` _RowSpacing_`, ` _ColumnSpacing_`, ` _RowHeight_`, ` _ColumnWidth_`, ` _AdjustForPageBreaks_`, ` _ShowSummaryTasks_`, ` _ViewBackgroundColor_`, ` _ViewBackgroundPattern_`, ` _ShowProgressMarks_`, ` _ShowPageBreaks_`, ` _ShowIDOnly_` )

 _expression_ An expression that returns an [Application](./Project.Application.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LayoutMode_|Optional|**Long**|Specifies whether the layout of boxes is controlled automatically or by the user, either with the  **LayoutNow** method or through the interface. Can be one of the **[PjLayoutMode](Project.PjLayoutMode.md)** constants.|
| _LayoutScheme_|Optional|**Long**|Specifies box alignment within each row. Can be one of the  **[PjLayoutScheme](Project.PjLayoutScheme.md)** constants.|
| _SummaryPrecedence_|Optional|**Boolean**|If  **True**, summary tasks are placed before subtasks.|
| _RowAlignment_|Optional|**Long**|Alignment of text within a row. Can be one of the  **[PjVerticalAlignment](Project.PjVerticalAlignment.md)** constants.|
| _ColumnAlignment_|Optional|**Long**|Alignment of text within a column. Can be one of the  **[PjAlignment](Project.PjAlignment.md)** constants.|
| _RowSpacing_|Optional|**Long**|Spacing between rows. The value can be from 0 to 200.|
| _ColumnSpacing_|Optional|**Long**| Spacing between columns. The value can be from 0 to 200.|
| _RowHeight_|Optional|**Long**|The height of each row of boxes. Can be one of the  **[PjRowColSize](Project.PjRowColSize.md)** constants.|
| _ColumnWidth_|Optional|**Long**|The width of each column of boxes. Can be one of the  **[PjRowColSize](Project.PjRowColSize.md)** constants.|
| _AdjustForPageBreaks_|Optional|**Boolean**|If  **True**, a new task is placed on the next page if it does not fit on the current page. If **False**, a new task can fall on a break between pages.|
| _ShowSummaryTasks_|Optional|**Boolean**|If  **True**, summary tasks are shown. If **False**, summary tasks are hidden.|
| _ViewBackgroundColor_|Optional|**Long**|The background color of the view. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _ViewBackgroundPattern_|Optional|**Long**|The pattern used for the background. Can be one of the  **[PjBackgroundPattern](Project.PjBackgroundPattern.md)** constants.|
| _ShowProgressMarks_|Optional|**Boolean**|**True** if tasks in progress are marked with a diagonal line from the upper-left corner of the box to the lower-right corner and completed tasks are marked with an additional diagonal line from the upper-right corner of the box to the lower-left corner. **False** if the progress of tasks is not marked.|
| _ShowPageBreaks_|Optional|**Boolean**|**True** if page breaks show in the Network Diagram; otherwise, **False**.|
| _ShowIDOnly_|Optional|**Boolean**|**True** if only task ID numbers are displayed. **False** if all the task data fields in Network Diagram boxes are displayed.|

### Return value

 **Boolean**


## Remarks

Using the  **BoxLayoutEx** method without specifying any arguments displays the **Box Layout** dialog box.


## Example

The following example sets the layout of boxes on the active Network Diagram view to the default values.


```vb
Sub ReturnToDefault()
    Application.BoxLayoutEx LayoutMode:=pjLayoutManual, LayoutScheme:=pjLayoutTopDownFromLeft, _
        SummaryPrecedence:=True, RowAlignment:=pjCenter, ColumnAlignment:=pjMiddle, RowSpacing:=45, _
        ColumnSpacing:=60, RowHeight:=pjSizeBestFit, ColumnWidth:=pjSizeBestFit, AdjustForPageBreaks:=True, _
        ShowSummaryTasks:=True, ViewBackgroundColor:=&HFFFFFF, ViewBackgroundPattern:=pjBackgroundSolidFill, _
        ShowProgressMarks:=False, ShowPageBreaks:=True, ShowIDOnly:=False
End Sub
```


 **Note**  If you use any of the  **PjColor** constants for the _ViewBackgroundColor_ parameter, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **BoxLayoutEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[BoxLayout](Project.Application.BoxLayout.md)** method.


