---
title: Create alternating row colors on a report
ms.prod: access
ms.assetid: ea37a0cb-9057-e268-28a7-183751c8a1b8
ms.date: 09/26/2018
localization_priority: Normal
---


# Create alternating row colors on a report

By default, Access formats each row of a report's detail section with the same background color. When printing a report, shading every other line of the detail section can make it much easier to read. You can use the **[AlternateBackColor](../../../api/Access.Section.AlternateBackColor.md)** property to specify a color to be displayed or printed on every other line in the detail section when viewing or printing a report.

The following example illustrates how to display light gray bars on every other line of the report's detail section when it is printed.

```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 
    Me.Section("Detail").AlternateBackColor = RGB(240, 240, 240) 
     
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
