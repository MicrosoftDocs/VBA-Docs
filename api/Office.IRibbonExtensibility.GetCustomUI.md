---
title: IRibbonExtensibility.GetCustomUI method (Office)
keywords: vbaof11.chm289001
f1_keywords:
- vbaof11.chm289001
ms.prod: office
api_name:
- Office.IRibbonExtensibility.GetCustomUI
ms.assetid: a0106415-999e-94da-379c-70fb7aa6119f
ms.date: 06/08/2017
localization_priority: Normal
---


# IRibbonExtensibility.GetCustomUI method (Office)

Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.


## Syntax

_expression_. `GetCustomUI`( `_RibbonID_` )

 _expression_ An expression that returns a [IRibbonExtensibility](Office.IRibbonExtensibility.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RibbonID_|Required|**String**|The ID for the RibbonX UI. |

## Return value

String


## Remarks

For Word, Excel, PowerPoint, and Access, there is only one ID for each application. Outlook uses ribbon extensibility to customize not only the ribbon in an inspector, but also the ribbon in an explorer, in various context menus, in contextual tabs in a ribbon, and in the Microsoft Office Backstage view. In each of these scenarios, the developer specifies the custom UI in an XML file that is loaded when Office calls  **GetCustomUI** with a specific ribbon ID.


## Example

In the following example, written in C#, the  **IRibbonExtensibility** interface is specified in the class definition. The example then implements the interfaces's only method, **GetCustomUI**. The method creates an instance of a **SteamReader** object that reads in the customization markup in an external XML file.


```vb
public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility 
... 
public string GetCustomUI(string RibbonID) 
{ 
 StreamReader customUIReader = new System.IO.StreamReader("C:\\RibbonXSampleCS\\customUI.xml"); 
 string customUIData = customUIReader.ReadToEnd(); 
 return customUIData; 
} 

```


## See also


[IRibbonExtensibility Object](Office.IRibbonExtensibility.md)



[IRibbonExtensibility Object Members](./overview/Library-Reference/iribbonextensibility-members-office.md)

