---
title: Windows object (Visio)
keywords: vis_sdr.chm10310
f1_keywords:
- vis_sdr.chm10310
ms.prod: visio
api_name:
- Visio.Windows
ms.assetid: 3fa64269-adde-3918-9970-3ce412d638f2
ms.date: 06/19/2019
localization_priority: Normal
---


# Windows object (Visio)

Includes a **[Window](Visio.Window.md)** object for a window that is open in the application.


## Remarks

To retrieve a **Windows** collection, use the **Windows** property of an **[Application](visio.application.windows.md)** object or a **[Window](Visio.Window.Windows.md)** object.

The default property of a **Windows** collection is **Item**.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this collection maps to the following types:

- **Microsoft.Office.Interop.Visio.IVWindows.GetEnumerator()** (to enumerate the **Window** objects)   
- **Microsoft.Office.Interop.Visio.IVWindows**

## Events

-  [BeforeWindowClosed](Visio.Windows.BeforeWindowClosed.md)
-  [BeforeWindowPageTurn](Visio.Windows.BeforeWindowPageTurn.md)
-  [BeforeWindowSelDelete](Visio.Windows.BeforeWindowSelDelete.md)
-  [KeyDown](Visio.Windows.KeyDown.md)
-  [KeyPress](Visio.Windows.KeyPress.md)
-  [KeyUp](Visio.Windows.KeyUp.md)
-  [MouseDown](Visio.Windows.MouseDown.md)
-  [MouseMove](Visio.Windows.MouseMove.md)
-  [MouseUp](Visio.Windows.MouseUp.md)
-  [OnKeystrokeMessageForAddon](Visio.Windows.OnKeystrokeMessageForAddon.md)
-  [QueryCancelWindowClose](Visio.Windows.QueryCancelWindowClose.md)
-  [SelectionChanged](Visio.Windows.SelectionChanged.md)
-  [ViewChanged](Visio.Windows.ViewChanged.md)
-  [WindowActivated](Visio.Windows.WindowActivated.md)
-  [WindowChanged](Visio.Windows.WindowChanged.md)
-  [WindowCloseCanceled](Visio.Windows.WindowCloseCanceled.md)
-  [WindowOpened](Visio.Windows.WindowOpened.md)
-  [WindowTurnedToPage](Visio.Windows.WindowTurnedToPage.md)

## Methods

-  [Add](Visio.Windows.Add.md)
-  [Arrange](Visio.Windows.Arrange.md)

## Properties

-  [Application](Visio.Windows.Application.md)
-  [Count](Visio.Windows.Count.md)
-  [EventList](Visio.Windows.EventList.md)
-  [Item](Visio.Windows.Item.md)
-  [ItemEx](Visio.Windows.ItemEx.md)
-  [ItemFromID](Visio.Windows.ItemFromID.md)
-  [ObjectType](Visio.Windows.ObjectType.md)
-  [PersistsEvents](Visio.Windows.PersistsEvents.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]