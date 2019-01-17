---
title: LastDLLError property (Visual Basic for Applications)
keywords: vblr6.chm1020108
f1_keywords:
- vblr6.chm1020108
ms.prod: office
api_name:
- Office.LastDLLError
ms.assetid: ed946e1e-a60a-576f-c6b6-0bec91b3d21d
ms.date: 12/19/2018
localization_priority: Normal
---


# LastDLLError property 

Returns a system error code produced by a call to a [dynamic-link library](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) (DLL). Read-only. **LastDLLError** always returns zero on the Macintosh.

## Remarks

The **LastDLLError** property applies only to DLL calls made from Visual Basic code. When such a call is made, the called function usually returns a code indicating success or failure, and the **LastDLLError** property is filled. 

Check the documentation for the DLL's functions to determine the return values that indicate success or failure. Whenever the failure code is returned, the Visual Basic application should immediately check the **LastDLLError** property. No exception is raised when the **LastDLLError** property is set.

## Example

When pasted into a **[UserForm](userform-window.md)** module, the following code causes an attempt to call a DLL function. The call fails because the argument that is passed in (a null pointer) generates an error, and in any event, SQL can't be cancelled if it isn't running. 

The code following the call checks the return of the call, and then prints at the **LastDLLError** property of the **[Err](err-object.md)** object to reveal the error code. On systems without DLLs, **LastDLLError** always returns zero.


```vb
Private Declare Function SQLCancel Lib "ODBC32.dll" _
 (ByVal hstmt As Long) As Integer

Private Sub UserForm_Click()
    Dim RetVal
    ' Call with invalid argument.
    RetVal = SQLCancel(myhandle&)
    ' Check for SQL error code.    
    If RetVal = -2 Then
        'Display the information code.
        MsgBox "Error code is :" & Err. LastDllError 
    End If
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]