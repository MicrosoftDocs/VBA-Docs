---
title: Report.PrintSection property (Access)
keywords: vbaac10.chm13730
f1_keywords:
- vbaac10.chm13730
ms.prod: access
api_name:
- Access.Report.PrintSection
ms.assetid: 745f4624-557b-0a4c-d4f4-9f0ba4113a61
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.PrintSection property (Access)

The **PrintSection** property specifies whether a section should be printed. Read/write **Boolean**.


## Syntax

_expression_.**PrintSection**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **PrintSection** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True**|(Default) The section is printed.|
|**False**|The section isn't printed.|

> [!NOTE] 
> To set this property, specify a macro or event procedure for a section's **[OnFormat](Access.Section.OnFormat.md)** property.

Microsoft Access sets this property to **True** before each section's **Format** event.


## Example

The following example does not print the section PageHeaderSection of the **Product Summary** report.

```vb
Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer) 
 
 Reports("Product Summary").PrintSection = False 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]