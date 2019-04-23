---
title: Report.MoveLayout property (Access)
keywords: vbaac10.chm13732
f1_keywords:
- vbaac10.chm13732
ms.prod: access
api_name:
- Access.Report.MoveLayout
ms.assetid: b02ddbda-ea3f-aad7-5f92-3b308dac4e79
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.MoveLayout property (Access)

The **MoveLayout** property specifies whether Microsoft Access should move to the next printing location on the page. Read/write **Boolean**.


## Syntax

_expression_.**MoveLayout**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **MoveLayout** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True**|(Default) The section's **Left** and **Top** properties are advanced to the next print location.|
|**False**|The section's **Left** and **Top** properties are unchanged.|

To set this property, specify an [event procedure](../access/Concepts/Settings/set-properties-by-using-visual-basic.md) for a section's **[OnFormat](Access.Section.OnFormat.md)** property.

Microsoft Access sets this property to **True** before each section's **Format** event.


## Example

The following example sets the **MoveLayout** property for the **Purchase Order** report to its default setting.

```vb
Reports("Purchase Order").MoveLayout = True 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]