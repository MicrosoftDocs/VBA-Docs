---
title: Dialogs.Item property (Excel)
keywords: vbaxl10.chm254074
f1_keywords:
- vbaxl10.chm254074
api_name:
- Excel.Dialogs.Item
ms.assetid: f9200ca3-711b-92ee-81b2-7c9cf1d104af
ms.date: 04/25/2019
ms.localizationpriority: medium
---


# Dialogs.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Dialogs](Excel.Dialogs.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **[XlBuiltInDialog](excel.xlbuiltindialog.md)** | **Variant**. The name or index number of the object.|

## Example

This example displays the **Open** dialog box and selects the **Read-Only** option.

```vb
Application.Dialogs.Item(xlDialogOpen).Show arg3:=True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]