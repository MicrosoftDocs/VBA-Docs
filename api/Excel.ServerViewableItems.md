---
title: ServerViewableItems object (Excel)
keywords: vbaxl10.chm832072
f1_keywords:
- vbaxl10.chm832072
ms.prod: excel
api_name:
- Excel.ServerViewableItems
ms.assetid: ce51dc80-ae34-f31a-81c0-f29467668289
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerViewableItems object (Excel)

A collection of objects that have been marked as viewable on the server.


## Remarks

This is a collection of references to objects in the workbook. Only objects in this collection will be shown on the server. By default, the entire workbook (including all worksheets) is shown.

Only one  **ServerViewableItems** object can exist per workbook. This collection is not indexable by name because there is no guarantee that the names of objects that are marked as viewable on the server are unique.

In the Excel user interface, you can view the collection of objects that are marked as viewable on the server in the  **Excel Services Options** dialog box.


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]