---
title: IVisEventProc object (Visio)
keywords: vis_sdr.chm60150
f1_keywords:
- vis_sdr.chm60150
ms.prod: visio
ms.assetid: 332ec60d-c70a-9d7f-15ad-bb797f60b3a5
ms.date: 06/19/2019
localization_priority: Normal
---


# IVisEventProc object (Visio)

The interface for handling event notifications in Microsoft Visio. 


## Remarks

In addition to the methods inherited from **IDispatch**, the **IVisEventProc** interface contains a single function, **VisEventProc**, which returns a **Variant**. Because **IVisEventProc** inherits from **IDispatch** and hence from **IUnknown**, you must implement the methods in those interfaces as well as the **VisEventProc** method.

To handle event notifications in Visio, create a class module that implements the **IVisEventProc** interface in Microsoft Visual Basic for Applications (VBA) or Microsoft Visual Basic, and then create an instance of this class to pass as an argument to the **[AddAdvise](Visio.EventList.AddAdvise.md)** method of the **EventList** collection.


## Methods

-  [VisEventProc](Visio.IVisEventProc.VisEventProc.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]