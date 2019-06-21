---
title: VisWebPageSettings.GetFormatName method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.GetFormatName
ms.assetid: 5586e07a-8b05-8894-d877-45c27584d4e0
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.GetFormatName method

Places the friendly name of the output format specified by the index passed to this method in the _pVal_ parameter passed to the method.


## Syntax

_expression_.**GetFormatName** (_nIndex_, _pVal_)

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_nIndex_ |Required| **Long**|The zero-based index of the output format within the list of output formats on the user's computer.|
|_pVal_ |Required| **String**|Variable that will hold the display name of the output format. The **GetFormatName** method places the name in this variable.|

## Return value

**Nothing**


## Remarks

You can view the available output formats in the **Save as Web Page** dialog box (**File** menu > **Save As Web Page** > **Publish** > **Advanced**).

You can determine the total number of formats by examining the **[FormatCount](Visio.VisWebPageSettings.FormatCount.md)** property. The formats include all those installed on the user's computer (for example, XAML, VML, JPG, GIF, PNG, and so on). To view a list of formats, see the topic for the **[ListFormats](Visio.VisWebPageSettings.ListFormats.md)** method.


## Example

The following example shows how to use the **GetFormatName** method to determine the display name of the output format being passed to the method.

```vb
Public Sub GetFormatName_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 Dim strName As String 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 vsoWebSettings.GetFormatName(1, strName) 
 
 Debug.Print strName 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]