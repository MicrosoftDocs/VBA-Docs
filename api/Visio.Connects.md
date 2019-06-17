---
title: Connects object (Visio)
keywords: vis_sdr.chm10070
f1_keywords:
- vis_sdr.chm10070
ms.prod: visio
api_name:
- Visio.Connects
ms.assetid: 8ac06fd8-0bbb-e9df-a08c-d697c4ac238e
ms.date: 06/08/2017
localization_priority: Normal
---


# Connects object (Visio)

 Includes a **Connect** object for each connection between two shapes in a drawing, such as a line and a box in an organization chart.


## Remarks

The default property of a  **Connects** collection is **Item**.

Use the  **Connects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object to which the indicated **Shape** object is connected (glued).

Use the  **FromConnects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object that is connected (glued) to the indicated **Shape** object.

Use the  **Connects** property of a **Page** object to retrieve a **Connects** collection with an entry for every connection on the **Page** object.

Use the  **Connects** property of a **Master** object to retrieve a **Connects** collection with an entry for every connection in the **Master** object.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVConnects.GetEnumerator()** (to enumerate the **Connect** objects.)
    

## Properties



|Name|
|:-----|
|[Application](Visio.Connects.Application.md)|
|[Count](Visio.Connects.Count.md)|
|[Document](Visio.Connects.Document.md)|
|[FromSheet](Visio.Connects.FromSheet.md)|
|[Item](Visio.Connects.Item.md)|
|[ObjectType](Visio.Connects.ObjectType.md)|
|[Stat](Visio.Connects.Stat.md)|
|[ToSheet](Visio.Connects.ToSheet.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]