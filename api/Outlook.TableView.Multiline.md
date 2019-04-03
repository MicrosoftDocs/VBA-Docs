---
title: TableView.MultiLine property (Outlook)
keywords: vbaol11.chm2524
f1_keywords:
- vbaol11.chm2524
ms.prod: outlook
api_name:
- Outlook.TableView.Multiline
ms.assetid: 732b39ca-ec7f-5a43-db55-3351a368b599
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.MultiLine property (Outlook)

Returns or sets an  **[OlMultiLine](Outlook.OlMultiLine.md)** constant that determines how multiple lines are displayed in the **[TableView](Outlook.TableView.md)** object. Read/write.


## Syntax

_expression_. `Multiline`

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Remarks

If the value of the  **[AutomaticColumnSizing](Outlook.TableView.AutomaticColumnSizing.md)** property is set to **False** or if the value of the **[AllowInCellEditing](Outlook.TableView.AllowInCellEditing.md)** property is set to **True**, the value of this property is automatically set to **olAlwaysSingleLine**.


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]