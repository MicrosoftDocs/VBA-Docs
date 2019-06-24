---
title: Application.RegisterRibbonX method (Visio)
keywords: vis_sdr.chm10062090
f1_keywords:
- vis_sdr.chm10062090
ms.prod: visio
api_name:
- Visio.Application.RegisterRibbonX
ms.assetid: 178db1c3-b3af-aa3f-af03-1aec1eab549a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RegisterRibbonX method (Visio)

Registers the **[IRibbonExtensibility](Office.IRibbonExtensibility.md)** interface that is implemented by the specified add-on to populate the custom user interface (UI).


## Syntax

_expression_.**RegisterRibbonX** (_SourceAddOn_, _TargetDocument_, _TargetModes_, _FriendlyName_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SourceAddOn_|Required| **IRibbonExtensibility**|The add-on to register.|
| _TargetDocument_|Required| **[Document](Visio.Document.md)**|The document that uses the custom UI.|
| _TargetModes_|Required| **[VisRibbonXModes](Visio.VisRibbonXModes.md)**|The modes in which the custom UI should be visible. See Remarks for possible values.|
| _FriendlyName_|Required| **String**|The name to associate with the UI items and errors that originate in the add-on.|

## Return value

 **Nothing**


## Remarks

If  _TargetDocument_ is null, the custom UI is defined at the level of the application. Otherwise, Microsoft Visio binds the visibility of the custom UI to the specified document. The UI does not appear in conjunction with any other document.

 _TargetModes_ can be one or more of the following **VisRibbonXModes** constants.



|Name|Value|Description|
|:-----|:-----|:-----|
| **visRxModeNone**|0|Display the custom UI when no documents are open.|
| **visRxModeDrawing**|1|Display the custom UI in Drawing mode.|
| **visRxModeStencil**|2|Display the custom UI in Stencil mode.|
| **visRxModePrintPreview**|4|Display the custom UI in Print Preview mode.|

If  _FriendlyName_ is null, the method fails.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]