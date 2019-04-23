---
title: Report.Page property (Access)
keywords: vbaac10.chm13721
f1_keywords:
- vbaac10.chm13721
ms.prod: access
api_name:
- Access.Report.Page
ms.assetid: 6d1dd330-ecd8-3b5c-c851-26bf7e431f98
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.Page property (Access)

The **Page** property specifies the current page number when a report is being printed. Read/write **Long**.


## Syntax

_expression_.**Page**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

Although you can set the **Page** property to a value, you most often use this property to return information about page numbers.

This property is only available in Print Preview or when printing.


## Example

The following example updates the report's caption to display the current position in the report as the user pages back and forth in the report.

```vb
Private Sub Report_Page()
    Me.Caption = "Now Viewing Page " & Me.Page & " Of " & Me.Pages & " Page(s)"
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]