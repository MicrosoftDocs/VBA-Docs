---
title: Write your own customized handler
ROBOTS: INDEX
ms.prod: access
ms.assetid: 67186df9-26b9-428d-2987-cd0bc165f231
ms.date: 06/08/2017
localization_priority: Normal
---


# Write your own customized handler

**Applies to:** Access 2013 | Access 2016

You may want to write your own handler if you are an IIS server administrator who wants the default RDS support, but more control over user requests and access rights.

The MSDFMAP.Handler implements the **IDataFactoryHandler** interface.

## IDataFactoryHandler interface

This interface has two methods, **GetRecordset** and **Reconnect**. Both methods require that the [CursorLocation](https://msdn.microsoft.com/library/8A048BD4-AE25-A555-1C07-14364B7E6560%28Office.15%29.aspx) property be set to **adUseClient**.

Both methods take arguments that appear after the first comma in the "**Handler=**" keyword. For example, `"Handler=progid,arg1,arg2;"` will pass an argument string of `"arg1,arg2"`, and will pass an argument string of  `"arg1,arg2"`, and  `"Handler=progid"` will pass a null argument.


## GetRecordset method

This method queries the data source and creates a new [Recordset](https://msdn.microsoft.com/library/0F963BF8-F066-DC8A-B754-F427DE712DF1%28Office.15%29.aspx) object using the arguments provided. The **Recordset** must be opened with **adLockBatchOptimistic** and must not be opened asynchronously.


### Arguments

_conn_ The connection string.

_args_ The arguments for the handler.

_query_ The command text for making a query.

_ppRS_ The pointer where the **Recordset** should be returned.


## Reconnect method

This method updates the data source. It creates a new [Connection](https://msdn.microsoft.com/library/C16023AA-0321-2513-EE71-255D6FFBA03D%28Office.15%29.aspx) object and attaches the given **Recordset**.


### Arguments

_conn_ The connection string.

_args_ The arguments for the handler.

_pRS_ A **Recordset** object.


### msdfhdl.idl

This is the interface definition for **IDataFactoryHandler** that appears in the **msdfhdl.idl** file.

```idl
[ 
  uuid(D80DE8B3-0001-11d1-91E6-00C04FBBBFB3), 
  version(1.0) 
] 
library MSDFHDL 
{ 
    importlib("stdole32.tlb"); 
    importlib("stdole2.tlb"); 
 
    // TLib : Microsoft ActiveX Data Objects 2.0 Library 
    // {00000200-0000-0010-8000-00AA006D2EA4} 
    #ifdef IMPLIB 
    importlib("implib\\x86\\release\\ado\\msado15.dll"); 
    #else 
    importlib("msado20.dll"); 
    #endif 
 
    [ 
      odl, 
      uuid(D80DE8B5-0001-11d1-91E6-00C04FBBBFB3), 
      version(1.0) 
    ] 
    interface IDataFactoryHandler : IUnknown 
    { 
HRESULT _stdcall GetRecordset( 
      [in] BSTR conn, 
      [in] BSTR args, 
      [in] BSTR query, 
      [out, retval] _Recordset **ppRS); 
 
// DataFactory will use the ActiveConnection property 
// on the Recordset after calling Reconnect. 
   HRESULT _stdcall Reconnect( 
      [in] BSTR conn, 
      [in] BSTR args, 
      [in] _Recordset *pRS); 
    }; 
}; 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]