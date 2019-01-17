---
title: Using events with embedded charts
keywords: vbaxl10.chm5205779
f1_keywords:
- vbaxl10.chm5205779
ms.prod: excel
ms.assetid: 1202370e-2e24-5b02-e52f-6f7c84facdd2
ms.date: 11/13/2018
localization_priority: Normal
---


# Using events with embedded charts

Events are enabled for chart sheets by default. Before you can use events with a **Chart** object that represents an embedded chart, you must create a new class module and declare an object of type **Chart** with events. For example, assume that a new class module is created and named EventClassModule. The new class module contains the following code.

```vb
Public WithEvents myChartClass As Chart
```

After the new object has been declared with events, it appears in the **Object** list box in the class module, and you can write event procedures for this object. (When you select the new object in the **Object** box, the valid events for that object are listed in the **Procedure** list box.)

Before your procedures will run, however, you must connect the declared object in the class module with the embedded chart. You can do this by using the following code from any module.

```vb
Dim myClassModule As New EventClassModule 
 
Sub InitializeChart() 
 Set myClassModule.myChartClass = _ 
 Charts(1).ChartObjects(1).Chart 
End Sub
```

After you run the InitializeChart procedure, the **myChartClass** object in the class module points to embedded chart 1 on worksheet 1, and the event procedures in the class module will run when the events occur.

## See also

- [Excel functions (by category)](https://support.office.com/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]