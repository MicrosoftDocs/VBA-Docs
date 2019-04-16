---
title: IconView.Copy method (Outlook)
keywords: vbaol11.chm2562
f1_keywords:
- vbaol11.chm2562
ms.prod: outlook
api_name:
- Outlook.IconView.Copy
ms.assetid: aa0c2905-766b-55d7-db32-07caffd03815
ms.date: 06/08/2017
localization_priority: Normal
---


# IconView.Copy method (Outlook)

Creates a new  **[View](Outlook.View.md)** object based on the existing **[IconView](Outlook.IconView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

_expression_ A variable that represents an [IconView](Outlook.IconView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## See also


[IconView Object](Outlook.IconView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]