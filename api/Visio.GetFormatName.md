---
title: VisWebPageSettings.GetFormatName Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.GetFormatName
ms.assetid: 5586e07a-8b05-8894-d877-45c27584d4e0
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings.GetFormatName Method (Visio Save As Web)

Places the friendly name of the output format specified by the index passed to this method in the pVal parameter passed to the method.


## Syntax

 _expression_. **GetFormatName**(**_nIndex_**,  **_pVal_**)

 _expression_An expression that returns a  ** [VisWebPageSettings](./overview/Visio.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|nIndex |Required| **Long**|The zero-based index of the output format within the list of output formats on the user's computer.|
|pVal |Required| **String**|Variable that will hold the display name of the output format. The  **GetFormatName** method places the name in this variable.|

## Return value

 **Nothing**


## Remarks

You can view the available output formats in the  **Save as Web Page** dialog box. (On the **File** menu, click **Save As Web Page**, click  **Publish**, and then click  **Advanced**).

You can determine the total number of formats by examining the  **[FormatCount](Visio.FormatCount.md)** property. The formats include all those installed on the user's computer (for example, XAML, VML, JPG, GIF, PNG, and so on). To view a list of formats, see the topic for the **[ListFormats](Visio.ListFormats.md)** method.


## Example

The following example shows how to use the  **GetFormatName** method to determine the display name of the output format being passed to the method.


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


