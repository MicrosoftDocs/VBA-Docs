---
title: WebOptions.LocationOfComponents property (Excel)
keywords: vbaxl10.chm662081
f1_keywords:
- vbaxl10.chm662081
ms.prod: excel
api_name:
- Excel.WebOptions.LocationOfComponents
ms.assetid: 0581343b-e93e-1413-4348-529f48a166eb
ms.date: 05/18/2019
localization_priority: Normal
---


# WebOptions.LocationOfComponents property (Excel)

Returns or sets the central URL (on the intranet or web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write **String**.


## Syntax

_expression_.**LocationOfComponents**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

Office Web components are automatically downloaded with the specified webpage if the **[DownloadComponents](Excel.WebOptions.DownloadComponents.md)** property is set to **True**, the components are not already installed, the path is valid and points to a location that contains the necessary components, and the user has a valid Microsoft Office license.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]