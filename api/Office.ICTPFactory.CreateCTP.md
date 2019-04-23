---
title: ICTPFactory.CreateCTP method (Office)
keywords: vbaof11.chm304001
f1_keywords:
- vbaof11.chm304001
ms.prod: office
api_name:
- Office.ICTPFactory.CreateCTP
ms.assetid: 17be1aa2-5045-2c89-151b-6f00d1bae6c1
ms.date: 01/16/2019
localization_priority: Normal
---


# ICTPFactory.CreateCTP method (Office)

Creates an instance of a custom task pane.


## Syntax

_expression_.**CreateCTP** (_CTPAxID_, _CTPTitle_, _CTPParentWindow_)

_expression_ An expression that returns an **[ICTPFactory](Office.ICTPFactory.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CTPAxID_|Required|**String**|The CLSID or ProgID of a Microsoft ActiveX object. |
| _CTPTitle_|Required|**String**|The title for the task pane.|
| _CTPParentWindow_|Optional|**Variant**|The window that hosts the task pane. If not present, the parent of the task pane is the ActiveWindow of the host application.|

## Return value

CustomTaskPane


## Example

The following example, written in C#, creates an instance of a **CustomTaskPane** object through the **ICustomTaskPaneConsumer** interface and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create task panes by using the **CreateCTP** method. 

Note that the example assumes that the task pane is part of a COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references an ActiveX control, SampleActiveX.myControl, which was created in a separate project.


```cs
public class Connect : Object, Extensibility.IDTExtensibility2, ICustomTaskPaneConsumer 
... 
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
} 
```


> [!NOTE] 
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also

- [ICTPFactory object members](overview/Library-Reference/ictpfactory-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]