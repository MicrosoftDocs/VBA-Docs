---
title: Report.HasData property (Access)
keywords: vbaac10.chm13726
f1_keywords:
- vbaac10.chm13726
ms.prod: access
api_name:
- Access.Report.HasData
ms.assetid: e8827477-6877-ec7a-63e5-7f4de972f0bb
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.HasData property (Access)

You can use the **HasData** property to determine if a report is bound to an empty recordset. Read/write **Long**.


## Syntax

_expression_.**HasData**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **HasData** property is set by Microsoft Access. The value of this property can be read only while printing or while in Print Preview.

The **HasData** property uses the following settings.

|Value|Description|
|:-----|:-----|
|1|The object has data.|
|0|The object doesn't have data.|
|1|The object is unbound.|

You can use this property to determine whether to hide a subreport that has no data. For example, the following expression hides the subreport control when its report has no data.

```vb
Me!SubReportControl.Visible = Me!SubReportControl.Report.HasData
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]