---
title: ProtectedViewWindow object (Excel)
keywords: vbaxl10.chm914072
f1_keywords:
- vbaxl10.chm914072
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow
ms.assetid: 6a32240c-c90b-c51a-6f8e-c3ff496b9855
ms.date: 03/30/2019
localization_priority: Normal
---


# ProtectedViewWindow object (Excel)

Represents a Protected View window.


## Remarks

A Protected View window is used to display a workbook from a potentially unsafe location. Unsafe locations are defined as the following:

- Files opened from the Internet.   
- Attachments opened from Outlook.    
- Files blocked by File Block Policy.   
- Files that fail Office file validation.   
- Files explicitly opened in **Protected View** by using the **Open in Protected View** command of the **Open** button in the **Open** dialog box.
    
Workbooks displayed in a Protected View window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and data connections. For more information about Protected View windows, see [What is Protected View?](https://support.office.com/en-us/article/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653?ocmsassetID=HA010355931&CTT=1&CorrelationId=ad189265-115e-4f59-bdf0-ee99038a5bb0&ui=en-US&rs=en-US&ad=US)

To return a single **ProtectedViewWindow** object from the **[ProtectedViewWindows](Excel.ProtectedViewWindows.md)** collection, use **ProtectedViewWindows** (_index_), where _index_ is the index number of the window that you want to open. 

You can also access the **ProtectedViewWindow** object that represents the active Protected View window by using the **[ActiveProtectedViewWindow](Excel.Application.ActiveProtectedViewWindow.md)** property of the **Application** object.

After you access a **ProtectedViewWindow** object, use the **[Workbook](Excel.ProtectedViewWindow.Workbook.md)** property to access the **[Workbook](Excel.Workbook.md)** object that represents the workbook file that is open in the Protected View window. Because a Protected View window is designed to protect the user from potentially malicious code, the operations that you can perform by using a **Workbook** object returned by a **ProtectedViewWindow** object will be limited. Operations that are not allowed will return an error.


## Example

The following code example accesses the **Workbook** object that represents the workbook that is open in the first Protected View window.

```vb
Dim wbProtected As Workbook 
 
If Application.ProtectedViewWindows.Count > 0 Then 
    Set wbProtected = Application.ProtectedViewWindows(1).Workbook 
End If 

```

## Methods

- [Activate](Excel.ProtectedViewWindow.Activate.md)
- [Close](Excel.ProtectedViewWindow.Close.md)
- [Edit](Excel.ProtectedViewWindow.Edit.md)

## Properties

- [Caption](Excel.ProtectedViewWindow.Caption.md)
- [EnableResize](Excel.ProtectedViewWindow.EnableResize.md)
- [Height](Excel.ProtectedViewWindow.Height.md)
- [Left](Excel.ProtectedViewWindow.Left.md)
- [SourceName](Excel.ProtectedViewWindow.SourceName.md)
- [SourcePath](Excel.ProtectedViewWindow.SourcePath.md)
- [Top](Excel.ProtectedViewWindow.Top.md)
- [Visible](Excel.ProtectedViewWindow.Visible.md)
- [Width](Excel.ProtectedViewWindow.Width.md)
- [WindowState](Excel.ProtectedViewWindow.WindowState.md)
- [Workbook](Excel.ProtectedViewWindow.Workbook.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
