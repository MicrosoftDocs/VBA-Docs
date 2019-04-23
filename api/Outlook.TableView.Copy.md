---
title: TableView.Copy method (Outlook)
keywords: vbaol11.chm2504
f1_keywords:
- vbaol11.chm2504
ms.prod: outlook
api_name:
- Outlook.TableView.Copy
ms.assetid: 985b5aaa-1f66-77e3-a035-3e2030318bf8
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.Copy method (Outlook)

Creates a new  **[View](Outlook.View.md)** object based on the existing **[TableView](Outlook.TableView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]