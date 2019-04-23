---
title: Report.NextRecord property (Access)
keywords: vbaac10.chm13731
f1_keywords:
- vbaac10.chm13731
ms.prod: access
api_name:
- Access.Report.NextRecord
ms.assetid: 771508ff-9a2d-6317-2b23-a1c0b012e7ba
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.NextRecord property (Access)

The **NextRecord** property specifies whether a section should advance to the next record. Read/write **Boolean**.


## Syntax

_expression_.**NextRecord**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **NextRecord** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True**|(Default) The section advances to the next record.|
|**False**|The section doesn't advance to the next record.|

To set this property, specify an event procedure for a section's **[OnFormat](Access.Section.OnFormat.md)** property.

Microsoft Access sets this property to **True** before each section's **Format** event.


## Example

The following example sets the **NextRecord** property to **False** for a given report.

```vb
Public Sub ChangeNextRecord(r As Report) 
 r.NextRecord = False 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]