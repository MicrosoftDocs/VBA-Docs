---
title: ListObject.Publish method (Excel)
keywords: vbaxl10.chm734074
f1_keywords:
- vbaxl10.chm734074
ms.prod: excel
api_name:
- Excel.ListObject.Publish
ms.assetid: 8b25819d-51c3-f505-8b9c-184355c48055
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Publish method (Excel)

Publishes the **ListObject** object to a server that is running Microsoft SharePoint Foundation.


## Syntax

_expression_.**Publish** (_Target_, _LinkSource_)

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **Variant**|Contains an array of **String** values, as described in the Remarks section.|
| _LinkSource_|Required| **Boolean**|See the Remarks section.|

## Return value

A **String** value that represents the URL of the published list on the SharePoint site.


## Remarks

The _Target_ parameter contains an array of **String** elements, as described in the following table.

|Element #|Contents|
|:-----|:-----|
|0|URL of SharePoint server|
|1|ListName (Display Name)|
|2|Description of the list. Optional.|

If the **ListObject** object is not currently linked to a list on a SharePoint site, setting _LinkSource_ to **True** creates a new list on the specified SharePoint site. 

If the **ListObject** object is currently linked to a list on a SharePoint site, setting _LinkSource_ to **True** replaces the existing link (you can only link the list to one SharePoint site). 

If the **ListObject** object is not currently linked, setting _LinkSource_ to **False** leaves the **ListObject** object unlinked. 

If the **ListObject** object is currently linked to a SharePoint site, setting _LinkSource_ to **False** keeps the **ListObject** object linked to the current SharePoint site.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]