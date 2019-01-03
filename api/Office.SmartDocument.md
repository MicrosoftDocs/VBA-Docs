---
title: SmartDocument object (Office)
keywords: vbaof11.chm262000
f1_keywords:
- vbaof11.chm262000
ms.prod: office
api_name:
- Office.SmartDocument
ms.assetid: b56a86eb-a031-d50b-905e-ef8b91914d61
ms.date: 06/08/2017
---


# SmartDocument object (Office)

The  **SmartDocument** property of the **Document** object in Microsoft Word and the **Workbook** object in Microsoft Excel returns a **SmartDocument** object.


## Remarks

Use the  **SmartDocument** object to manage the XML expansion pack attached to the active document.

Use the  **SmartDocument** object's **SolutionID** and **SolutionURI** properties to retrieve information about the XML expansion pack attached to the active document or workbook. Use the **PickSolution** method to allow the user to select an available XML expansion pack from a list to attach to the active document or workbook. Use the **RefreshPane** method to refresh the smart document's **Document Actions** task pane.

The  **SmartDocument** object model is available whether or not a document has an XML expansion pack attached. The **SmartDocument** property of the **Document** or **Workbook** objects does not return **Nothing** when the active document has no XML expansion pack attached. Examine the **SolutionID** property to determine whether the active document has an XML expansion pack attached.


## Methods



|**Name**|
|:-----|
|[PickSolution](Office.SmartDocument.PickSolution.md)|
|[RefreshPane](Office.SmartDocument.RefreshPane.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.SmartDocument.Application.md)|
|[Creator](Office.SmartDocument.Creator.md)|
|[SolutionID](Office.SmartDocument.SolutionID.md)|
|[SolutionURL](Office.SmartDocument.SolutionURL.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
