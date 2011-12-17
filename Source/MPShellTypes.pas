unit MPShellTypes;

// Version 2.1.0
//
// The contents of this file are subject to the Mozilla Public License
// Version 1.1 (the "License"); you may not use this file except in compliance
// with the License. You may obtain a copy of the License at http://www.mozilla.org/MPL/
//
// Alternatively, you may redistribute this library, use and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation;
// either version 2.1 of the License, or (at your option) any later version.
// You may obtain a copy of the LGPL at http://www.gnu.org/copyleft/.
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the
// specific language governing rights and limitations under the License.
//
// The initial developer of this code is Jim Kueneman <jimdk@mindspring.com>
// Special thanks to the following in no particular order for their help/support/code
//    Danijel Malik, Robert Lee, Werner Lehmann, Alexey Torgashin, Milan Vandrovec
//
// Portions provided by Derek Moore of Alastria Software
//
//----------------------------------------------------------------------------

interface

{$I Compilers.inc}
{$I Options.inc}
{$I ..\Include\Debug.inc}
{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  ImgList, ShlObj, ShellAPI, ActiveX, ComObj, CommCtrl;


{$HPPEMIT '#include <comdef.h>'}
{$HPPEMIT '#include <oleidl.h>'}
{$HPPEMIT '#include <comcat.h>'}
{$HPPEMIT '#include <comcat.h>'}

{$IFDEF CPPB_6}
  {$HPPEMIT '#include <shtypes.h>'} // if BCB6 or later
{$ENDIF}
{$HPPEMIT '#include <winioctl.h>'}

(*$HPPEMIT 'namespace Mpshellutilities { class TNamespace; }' *)
{$HPPEMIT 'typedef DelphiInterface<IDropTarget>   _di_IDropTarget;'}
{$HPPEMIT 'typedef DelphiInterface<IQueryInfo>    _di_IQueryInfo;'}
{$HPPEMIT 'typedef DelphiInterface<IEnumString> _di_IEnumString;'}
{$HPPEMIT 'typedef DelphiInterface<IBindCtx> _di_IBindCtx;'}
{$HPPEMIT 'typedef DelphiInterface<IClassFactory> _di_IClassFactory;'}
{$HPPEMIT 'typedef DelphiInterface<IDeskBand> _di_IDeskBand;'}
{$HPPEMIT 'typedef DelphiInterface<IDropSource> _di_IDropSource;'}

{$HPPEMIT 'typedef _SHELLDETAILS tagSHELLDETAILS;'}
(*$HPPEMIT 'namespace Activex { typedef System::DelphiInterface<IEnumGUID> _di_IEnumGUID; }' *)


//------------------------------------------------------------------------------
// Missing Windows Message definitions
//------------------------------------------------------------------------------

{$IFDEF COMPILER_4}
type
  TWMContextMenu = packed record
    Msg: Cardinal;
    hWnd: HWND;
    case Integer of
      0: (
        XPos: Smallint;
        YPos: Smallint);
      1: (
        Pos: TSmallPoint;
        Result: Longint);
  end;
{$ENDIF}

{$IFNDEF DELPHI_7_UP}
type
  TWMPrint = packed record
    Msg: Cardinal;
    DC: HDC;
    Flags: Cardinal;
    Result: Integer;
  end;

  TWMPrintClient = TWMPrint;
{$ENDIF}

const
  {$EXTERNALSYM SEE_MASK_WAITFORINPUTIDLE}
  SEE_MASK_WAITFORINPUTIDLE = $02000000;

const
  ComCtl_470 = $00040046;
  ComCtl_471 = $00040047;
  ComCtl_472 = $00040048;
  ComCtl_580 = $00050050;
  ComCtl_581 = $00050050;
  ComCtl_600 = $00060000;

const
  {$EXTERNALSYM FOF_NO_CONNECTED_ELEMENTS}
  FOF_NO_CONNECTED_ELEMENTS = $0000;
  {$EXTERNALSYM FOF_NOCOPYSECURITYATTRIBS}
  FOF_NOCOPYSECURITYATTRIBS = $0000;
  {$EXTERNALSYM FOF_NORECURSION}
  FOF_NORECURSION = $0000;
  {$EXTERNALSYM FOF_NORECURSEREPARSE}
  FOF_NORECURSEREPARSE = $0000;
  {$EXTERNALSYM FOF_WANTNUKEWARNING}
  FOF_WANTNUKEWARNING = $0000;

  {$EXTERNALSYM GIL_DEFAULTICON}
  GIL_DEFAULTICON =  $0040;      // get the default icon location if the final one takes too long to get
  {$EXTERNALSYM GIL_FORSHORTCUT}
  GIL_FORSHORTCUT =  $0080;      // the icon is for a shortcut to the object

  SID_IShellIconOverlayIdentifier = '{0C6C4200-C589-11D0-999A-00C04FD655E1}';
  IID_IShellIconOverlayIdentifier: TGUID = SID_IShellIconOverlayIdentifier;

  {$IFNDEF COMPILER_14_UP}
  SID_IShellLibrary = '{11A66EFA-382E-451A-9234-1E0E12EF3085}';
  IID_IShellLibrary: TGUID = SID_IShellLibrary;
  {$ENDIF}

  {$EXTERNALSYM ISIOI_ICONFILE}
  ISIOI_ICONFILE =  $00000001;         // path is returned through pwszIconFile
  {$EXTERNALSYM ISIOI_ICONINDEX}
  ISIOI_ICONINDEX = $00000002;          // icon index in pwszIconFile is returned through pIndex

type
  IShellIconOverlayIdentifier = interface(IUnknown)
  [SID_IShellIconOverlayIdentifier]
    function IsMemberOf(pwszPath: LPCWSTR; dwAttrib: DWORD): HRESULT; stdcall;
    function GetOverlayInfo(pwszIconFile: LPWSTR; cchMax: Integer; var pIndex: Integer; var pdwFlags: DWORD): HRESULT; stdcall;
    function GetPriority(var pPriority: Integer): HRESULT; stdcall;
  end;

const
  {$EXTERNALSYM FILE_ATTRIBUTE_ENCRYPTED}
  FILE_ATTRIBUTE_ENCRYPTED     = $00004000;
  {$EXTERNALSYM FILE_ATTRIBUTE_REPARSE_POINT}
  FILE_ATTRIBUTE_REPARSE_POINT   = $00000400;
  {$EXTERNALSYM FILE_ATTRIBUTE_SPARSE_FILE}
  FILE_ATTRIBUTE_SPARSE_FILE     = $00000200;
  {$EXTERNALSYM FILE_ATTRIBUTE_NOT_CONTENT_INDEXED}
  FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = $00002000;


  {$EXTERNALSYM IO_REPARSE_TAG_DFS}
  IO_REPARSE_TAG_DFS = $8000000A;
  {$EXTERNALSYM IO_REPARSE_TAG_DFSR}
  IO_REPARSE_TAG_DFSR = $80000012;
  {$EXTERNALSYM IO_REPARSE_TAG_HSM}
  IO_REPARSE_TAG_HSM = $C0000004;
  {$EXTERNALSYM IO_REPARSE_TAG_HSM2}
  IO_REPARSE_TAG_HSM2 = $80000006;
  {$EXTERNALSYM IO_REPARSE_TAG_MOUNT_POINT}
  IO_REPARSE_TAG_MOUNT_POINT = $A0000003;
  {$EXTERNALSYM IO_REPARSE_TAG_SIS}
  IO_REPARSE_TAG_SIS = $80000007;
  {$EXTERNALSYM IO_REPARSE_TAG_SYMLINK}
  IO_REPARSE_TAG_SYMLINK = $A000000C;

  {$EXTERNALSYM FILE_FLAG_OPEN_REPARSE_POINT}
  FILE_FLAG_OPEN_REPARSE_POINT = $00200000;

  {$EXTERNALSYM FSCTL_SET_REPARSE_POINT}
  FSCTL_SET_REPARSE_POINT = $000900A4;
  {$EXTERNALSYM FSCTL_GET_REPARSE_POINT}
  FSCTL_GET_REPARSE_POINT = $000900A8;
  {$EXTERNALSYM FSCTL_DELETE_REPARSE_POINT}
  FSCTL_DELETE_REPARSE_POINT = $000900AC;
  {$EXTERNALSYM FSCTL_CREATE_OR_GET_OBJECT_ID}
  FSCTL_CREATE_OR_GET_OBJECT_ID = $000900C0;
  {$EXTERNALSYM FSCTL_SET_SPARSE}
  FSCTL_SET_SPARSE = $000900C4;
  {$EXTERNALSYM FSCTL_SET_ZERO_DATA}
  FSCTL_SET_ZERO_DATA = $000900C8;
  {$EXTERNALSYM FSCTL_SET_ENCRYPTION}
  FSCTL_SET_ENCRYPTION = $000900D7;

(*$HPPEMIT 'typedef struct _REPARSE_DATA_BUFFER {'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '    DWORD   ReparseTag;'*)
(*$HPPEMIT '    WORD    ReparseDataLength;'*)
(*$HPPEMIT '    WORD    Reserved;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '    union {'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '        struct {'*)
(*$HPPEMIT '            WORD    SubstituteNameOffset;'*)
(*$HPPEMIT '            WORD    SubstituteNameLength;'*)
(*$HPPEMIT '            WORD    PrintNameOffset;'*)
(*$HPPEMIT '            WORD    PrintNameLength;'*)
(*$HPPEMIT '            WCHAR   PathBuffer[1];'*)
(*$HPPEMIT '        } SymbolicLinkReparseBuffer;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '        struct {'*)
(*$HPPEMIT '            WORD    SubstituteNameOffset;'*)
(*$HPPEMIT '            WORD    SubstituteNameLength;'*)
(*$HPPEMIT '            WORD    PrintNameOffset;'*)
(*$HPPEMIT '            WORD    PrintNameLength;'*)
(*$HPPEMIT '            WCHAR   PathBuffer[1];'*)
(*$HPPEMIT '        } MountPointReparseBuffer;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '        struct {'*)
(*$HPPEMIT '            UCHAR   DataBuffer[1];'*)
(*$HPPEMIT '        } GenericReparseBuffer;'*)
(*$HPPEMIT '    };'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '} REPARSE_DATA_BUFFER, *PREPARSE_DATA_BUFFER;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '#ifndef REPARSE_DATA_BUFFER_HEADER_SIZE'*)
(*$HPPEMIT '#define REPARSE_DATA_BUFFER_HEADER_SIZE   8'*)
(*$HPPEMIT '#endif'*)
(*$HPPEMIT ''*)
(*$HPPEMIT 'typedef struct _REPARSE_POINT_INFORMATION {'*)
(*$HPPEMIT '        WORD    ReparseDataLength;'*)
(*$HPPEMIT '        WORD    UnparsedNameLength;'*)
(*$HPPEMIT '} REPARSE_POINT_INFORMATION, *PREPARSE_POINT_INFORMATION;'*)
(*$HPPEMIT ''*)
(*$HPPEMIT '#ifndef IO_REPARSE_TAG_VALID_VALUES'*)
(*$HPPEMIT '#define IO_REPARSE_TAG_VALID_VALUES 0x0E000FFFF'*)
(*$HPPEMIT '#endif'*)
(*$HPPEMIT ''*)

type
  {$EXTERNALSYM _REPARSE_DATA_BUFFER}
  _REPARSE_DATA_BUFFER = record
    ReparseTag: DWORD;
    ReparseDataLength: Word;
    Reserved: Word;
    case Integer of
      0: ( // SymbolicLinkReparseBuffer and MountPointReparseBuffer
        SubstituteNameOffset: Word;
        SubstituteNameLength: Word;
        PrintNameOffset: Word;
        PrintNameLength: Word;
        PathBuffer: array [0..0] of WCHAR);
      1: ( // GenericReparseBuffer
        DataBuffer: array [0..0] of Byte);
  end;
  {$EXTERNALSYM REPARSE_DATA_BUFFER}
  REPARSE_DATA_BUFFER = _REPARSE_DATA_BUFFER;
  {$EXTERNALSYM PREPARSE_DATA_BUFFER}
  PREPARSE_DATA_BUFFER = ^_REPARSE_DATA_BUFFER;
  TReparseDataBuffer = _REPARSE_DATA_BUFFER;
  PReparseDataBuffer = PREPARSE_DATA_BUFFER;

const
  {$EXTERNALSYM REPARSE_DATA_BUFFER_HEADER_SIZE}
  REPARSE_DATA_BUFFER_HEADER_SIZE = 8;
  {$EXTERNALSYM MAXIMUM_REPARSE_DATA_BUFFER_SIZE}
  MAXIMUM_REPARSE_DATA_BUFFER_SIZE = DWORD(16 * 1024);

  {$EXTERNALSYM SYMLINK_FLAG_RELATIVE}
  SYMLINK_FLAG_RELATIVE = $00000001;




//------------------------------------------------------------------------------
// Property Sheets
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM PSN_TRANSLATEACCELERATOR}
  PSN_TRANSLATEACCELERATOR =  PSN_FIRST - 12;
  {$EXTERNALSYM PSN_QUERYINITIALFOCUS}
  PSN_QUERYINITIALFOCUS    =  PSN_FIRST - 13;

  {$EXTERNALSYM PSM_GETCURRENTPAGEHWND}
  PSM_GETCURRENTPAGEHWND = WM_USER + 118;

//------------------------------------------------------------------------------
// IRunnableTask
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM SID_IRunnableTask}
  SID_IRunnableTask = '{85788D00-6807-11d0-B810-00C04FD706EC}';

type
  {$EXTERNALSYM IRunnableTask}
  IRunnableTask = interface(IUnknown)
  [SID_IRunnableTask]
    function Run: HRESULT; stdcall;
    function Kill(fWait: BOOL): HRESULT; stdcall;
    function Suspend: HRESULT; stdcall;
    function Resume: HRESULT; stdcall;
    function IsRunning: ULONG; stdcall;       // Returns one of the IRTIR_TASK_xxxx constants
  end;

const
 {$EXTERNALSYM IRTIR_TASK_NOT_RUNNING}
  IRTIR_TASK_NOT_RUNNING =   0; // Extraction has not yet started.
  {$EXTERNALSYM IRTIR_TASK_RUNNING}
  IRTIR_TASK_RUNNING     =   1; // The task is running.
  {$EXTERNALSYM IRTIR_TASK_SUSPENDED}
  IRTIR_TASK_SUSPENDED   =   2; // The task is suspended.
  {$EXTERNALSYM IRTIR_TASK_PENDING}
  IRTIR_TASK_PENDING     =   3; // The thread has been killed but has not completely shut down yet.
  {$EXTERNALSYM IRTIR_TASK_FINISHED}
  IRTIR_TASK_FINISHED    =   4; // The task is finished.

 //------------------------------------------------------------------------------
// IQueryInfo
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM QITIPF_DEFAULT}
  QITIPF_DEFAULT         = $00000000;  // Normal Tip
  {$EXTERNALSYM QITIPF_USENAME}
  QITIPF_USENAME         = $00000001;  // Provide the name of the item in ppwszTip rather than the info tip text.
  {$EXTERNALSYM QITIPF_LINKNOTARGET}
  QITIPF_LINKNOTARGET    = $00000002;  // If the item is a shortcut, retrieve the info tip text of the shortcut rather than its target.
  {$EXTERNALSYM QITIPF_LINKUSETARGET}
  QITIPF_LINKUSETARGET   = $00000004;  // If the item is a shortcut, retrieve the info tip text of the shortcut's target.
  {$EXTERNALSYM QITIPF_USESLOWTIP}
  QITIPF_USESLOWTIP      = $00000008;  // Flag says it's OK to take a long time generating tip. Search the entire namespace for the information. This may result in a delayed response time.


//------------------------------------------------------------------------------
// Network Types
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM WNNC_NET_MSNET}
  WNNC_NET_MSNET       = $00010000;
  {$EXTERNALSYM WNNC_NET_LANMAN}
  WNNC_NET_LANMAN      = $00020000;
  {$EXTERNALSYM WNNC_NET_NETWARE}
  WNNC_NET_NETWARE     = $00030000;
  {$EXTERNALSYM WNNC_NET_VINES}
  WNNC_NET_VINES       = $00040000;
  {$EXTERNALSYM WNNC_NET_10NET}
  WNNC_NET_10NET       = $00050000;
  {$EXTERNALSYM WNNC_NET_LOCUS}
  WNNC_NET_LOCUS       = $00060000;
  {$EXTERNALSYM WNNC_NET_SUN_PC_NFS}
  WNNC_NET_SUN_PC_NFS  = $00070000;
  {$EXTERNALSYM WNNC_NET_LANSTEP}
  WNNC_NET_LANSTEP     = $00080000;
  {$EXTERNALSYM WNNC_NET_9TILES}
  WNNC_NET_9TILES      = $00090000;
  {$EXTERNALSYM WNNC_NET_LANTASTIC}
  WNNC_NET_LANTASTIC   = $000A0000;
  {$EXTERNALSYM WNNC_NET_AS400}
  WNNC_NET_AS400       = $000B0000;
  {$EXTERNALSYM WNNC_NET_FTP_NFS}
  WNNC_NET_FTP_NFS     = $000C0000;
  {$EXTERNALSYM WNNC_NET_PATHWORKS}
  WNNC_NET_PATHWORKS   = $000D0000;
  {$EXTERNALSYM WNNC_NET_LIFENET}
  WNNC_NET_LIFENET     = $000E0000;
  {$EXTERNALSYM WNNC_NET_POWERLAN}
  WNNC_NET_POWERLAN    = $000F0000;
  {$EXTERNALSYM WNNC_NET_BWNFS}
  WNNC_NET_BWNFS       = $00100000;
  {$EXTERNALSYM WNNC_NET_COGENT}
  WNNC_NET_COGENT      = $00110000;
  {$EXTERNALSYM WNNC_NET_FARALLON}
  WNNC_NET_FARALLON    = $00120000;
  {$EXTERNALSYM WNNC_NET_APPLETALK}
  WNNC_NET_APPLETALK   = $00130000;
  {$EXTERNALSYM WNNC_NET_INTERGRAPH}
  WNNC_NET_INTERGRAPH  = $00140000;
  {$EXTERNALSYM WNNC_NET_SYMFONET}
  WNNC_NET_SYMFONET    = $00150000;
   {$EXTERNALSYM WNNC_NET_CLEARCASE}
  WNNC_NET_CLEARCASE   = $00160000;
  {$EXTERNALSYM WNNC_NET_FRONTIER}
  WNNC_NET_FRONTIER    = $00170000;
  {$EXTERNALSYM WNNC_NET_BMC}
  WNNC_NET_BMC         = $00180000;
  {$EXTERNALSYM WNNC_NET_DCE}
  WNNC_NET_DCE         = $00190000;
  {$EXTERNALSYM WNNC_NET_AVID}
  WNNC_NET_AVID        = $001A0000;
  {$EXTERNALSYM WNNC_NET_DOCUSPACE}
  WNNC_NET_DOCUSPACE   = $001B0000;
  {$EXTERNALSYM WNNC_NET_MANGOSOFT}
  WNNC_NET_MANGOSOFT   = $001C0000;
  {$EXTERNALSYM WNNC_NET_SERNET}
  WNNC_NET_SERNET      = $001D0000;
  {$EXTERNALSYM WNNC_NET_RIVERFRONT1}
  WNNC_NET_RIVERFRONT1 = $001E0000;
  {$EXTERNALSYM WNNC_NET_RIVERFRONT2}
  WNNC_NET_RIVERFRONT2 = $001F0000;
  {$EXTERNALSYM WNNC_NET_DECORB}
  WNNC_NET_DECORB      = $00200000;
  {$EXTERNALSYM WNNC_NET_PROTSTOR}
  WNNC_NET_PROTSTOR    = $00210000;
  {$EXTERNALSYM WNNC_NET_FJ_REDIR}
  WNNC_NET_FJ_REDIR    = $00220000;
  {$EXTERNALSYM WNNC_NET_DISTINCT}
  WNNC_NET_DISTINCT    = $00230000;
  {$EXTERNALSYM WNNC_NET_TWINS}
  WNNC_NET_TWINS       = $00240000;
  {$EXTERNALSYM WNNC_NET_RDR2SAMPLE}
  WNNC_NET_RDR2SAMPLE  = $00250000;
  {$EXTERNALSYM WNNC_NET_CSC}
  WNNC_NET_CSC         = $00260000;
  {$EXTERNALSYM WNNC_NET_3IN1}
  WNNC_NET_3IN1        = $00270000;
  {$EXTERNALSYM WNNC_NET_EXTENDNET}
  WNNC_NET_EXTENDNET   = $00290000;
  {$EXTERNALSYM WNNC_NET_STAC}
  WNNC_NET_STAC        = $002A0000;
  {$EXTERNALSYM WNNC_NET_FOXBAT}
  WNNC_NET_FOXBAT      = $002B0000;
  {$EXTERNALSYM WNNC_NET_YAHOO}
  WNNC_NET_YAHOO       = $002C0000;
  {$EXTERNALSYM WNNC_NET_EXIFS}
  WNNC_NET_EXIFS       = $002D0000;
  {$EXTERNALSYM WNNC_NET_DAV}
  WNNC_NET_DAV         = $002E0000;
  {$EXTERNALSYM WNNC_NET_KNOWARE}
  WNNC_NET_KNOWARE     = $002F0000;
  {$EXTERNALSYM WNNC_NET_OBJECT_DIRE}
  WNNC_NET_OBJECT_DIRE = $00300000;
  {$EXTERNALSYM WNNC_NET_MASFAX}
  WNNC_NET_MASFAX      = $00310000;
  {$EXTERNALSYM WNNC_NET_HOB_NFS}
  WNNC_NET_HOB_NFS     = $00320000;
  {$EXTERNALSYM WNNC_NET_SHIVA}
  WNNC_NET_SHIVA       = $00330000;
  {$EXTERNALSYM WNNC_NET_IBMAL}
  WNNC_NET_IBMAL       = $00340000;
  {$EXTERNALSYM WNNC_NET_LOCK}
  WNNC_NET_LOCK        = $00350000;
  {$EXTERNALSYM WNNC_NET_TERMSRV}
  WNNC_NET_TERMSRV     = $00360000;
  {$EXTERNALSYM WNNC_NET_SRT}
  WNNC_NET_SRT         = $00370000;
  {$EXTERNALSYM WNNC_NET_QUINCY}
  WNNC_NET_QUINCY      = $00380000;


  NETWORK_PROVIDER_TYPES: array[0..54] of LongWord = (
    $00010000,
    $00020000,
    $00030000,
    $00040000,
    $00050000,
    $00060000,
    $00070000,
    $00080000,
    $00090000,
    $000A0000,
    $000B0000,
    $000C0000,
    $000D0000,
    $000E0000,
    $000F0000,
    $00100000,
    $00110000,
    $00120000,
    $00130000,
    $00140000,
    $00150000,
    $00160000,
    $00170000,
    $00180000,
    $00190000,
    $001A0000,
    $001B0000,
    $001C0000,
    $001D0000,
    $001E0000,
    $001F0000,
    $00200000,
    $00210000,
    $00220000,
    $00230000,
    $00240000,
    $00250000,
    $00260000,
    $00270000,
    $00290000,
    $002A0000,
    $002B0000,
    $002C0000,
    $002D0000,
    $002E0000,
    $002F0000,
    $00300000,
    $00310000,
    $00320000,
    $00330000,
    $00340000,
    $00350000,
    $00360000,
    $00370000,
    $00380000
  );

  NETWORK_PROVIDER_TYPE_IDS: array[0..54] of string = (
  'WNNC_NET_MSNET',
  'WNNC_NET_LANMAN',
  'WNNC_NET_NETWARE',
  'WNNC_NET_VINES',
  'WNNC_NET_10NET',
  'WNNC_NET_LOCUS',
  'WNNC_NET_SUN_PC_NFS',
  'WNNC_NET_LANSTEP',
  'WNNC_NET_9TILES',
  'WNNC_NET_LANTASTIC',
  'WNNC_NET_AS400',
  'WNNC_NET_FTP_NFS',
  'WNNC_NET_PATHWORKS',
  'WNNC_NET_LIFENET',
  'WNNC_NET_POWERLAN',
  'WNNC_NET_BWNFS',
  'WNNC_NET_COGENT',
  'WNNC_NET_FARALLON',
  'WNNC_NET_APPLETALK',
  'WNNC_NET_INTERGRAPH',
  'WNNC_NET_SYMFONET',
  'WNNC_NET_CLEARCASE',
  'WNNC_NET_FRONTIER ',
  'WNNC_NET_BMC',
  'WNNC_NET_DCE',
  'WNNC_NET_AVID',
  'WNNC_NET_DOCUSPACE',
  'WNNC_NET_MANGOSOFT',
  'WNNC_NET_SERNET',
  'WNNC_NET_RIVERFRONT1',
  'WNNC_NET_RIVERFRONT2',
  'WNNC_NET_DECORB',
  'WNNC_NET_PROTSTOR',
  'WNNC_NET_FJ_REDIR',
  'WNNC_NET_DISTINCT',
  'WNNC_NET_TWINS',
  'WNNC_NET_RDR2SAMPLE',
  'WNNC_NET_CSC',
  'WNNC_NET_3IN1',
  'WNNC_NET_EXTENDNET',
  'WNNC_NET_STAC',
  'WNNC_NET_FOXBAT',
  'WNNC_NET_YAHOO',
  'WNNC_NET_EXIFS',
  'WNNC_NET_DAV',
  'WNNC_NET_KNOWARE',
  'WNNC_NET_OBJECT_DIRE',
  'WNNC_NET_MASFAX',
  'WNNC_NET_HOB_NFS',
  'WNNC_NET_SHIVA',
  'WNNC_NET_IBMAL',
  'WNNC_NET_LOCK',
  'WNNC_NET_TERMSRV',
  'WNNC_NET_SRT',
  'WNNC_NET_QUINCY'
  );

{$IFNDEF COMPILER_7_UP}
{$EXTERNALSYM TBSTYLE_BUTTON}
  TBSTYLE_BUTTON          = $00;
  {$EXTERNALSYM TBSTYLE_SEP}
  TBSTYLE_SEP             = $01;
  {$EXTERNALSYM TBSTYLE_CHECK}
  TBSTYLE_CHECK           = $02;
  {$EXTERNALSYM TBSTYLE_GROUP}
  TBSTYLE_GROUP           = $04;
  {$EXTERNALSYM TBSTYLE_CHECKGROUP}
  TBSTYLE_CHECKGROUP      = TBSTYLE_GROUP or TBSTYLE_CHECK;
  {$EXTERNALSYM TBSTYLE_DROPDOWN}
  TBSTYLE_DROPDOWN        = $08;
  {$EXTERNALSYM TBSTYLE_AUTOSIZE}
  TBSTYLE_AUTOSIZE        = $0010;
  {$EXTERNALSYM TBSTYLE_NOPREFIX}
  TBSTYLE_NOPREFIX        = $0020;
  {$EXTERNALSYM BTNS_BUTTON}
  BTNS_BUTTON             = TBSTYLE_BUTTON;
  {$EXTERNALSYM BTNS_SEP}
  BTNS_SEP                = TBSTYLE_SEP;
  {$EXTERNALSYM BTNS_CHECK}
  BTNS_CHECK              = TBSTYLE_CHECK;
  {$EXTERNALSYM BTNS_GROUP}
  BTNS_GROUP              = TBSTYLE_GROUP;
  {$EXTERNALSYM BTNS_CHECKGROUP}
  BTNS_CHECKGROUP         = TBSTYLE_CHECKGROUP;
  {$EXTERNALSYM BTNS_DROPDOWN}
  BTNS_DROPDOWN           = TBSTYLE_DROPDOWN;
  {$EXTERNALSYM BTNS_AUTOSIZE}
  BTNS_AUTOSIZE           = TBSTYLE_AUTOSIZE;
  {$EXTERNALSYM BTNS_NOPREFIX}
  BTNS_NOPREFIX           = TBSTYLE_NOPREFIX;
  {$EXTERNALSYM BTNS_SHOWTEXT}
  BTNS_SHOWTEXT           = $0040;
  {$EXTERNALSYM BTNS_WHOLEDROPDOWN}
  BTNS_WHOLEDROPDOWN      = $0080;
{$ENDIF}

//------------------------------------------------------------------------------
// New ImageList styles
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM ILD_PRESERVEALPHA}
  ILD_PRESERVEALPHA = $00001000;

//------------------------------------------------------------------------------
// Some new magic extended Listview Styles
//------------------------------------------------------------------------------
{$ifndef COMPILER_11_UP}
const
  {$EXTERNALSYM LVS_EX_DOUBLEBUFFER}
  LVS_EX_DOUBLEBUFFER = $00010000;
{$endif}

//------------------------------------------------------------------------------
// Undocumented SHChangeNotifier Registration Constants and Types
//------------------------------------------------------------------------------
const
  {$EXTERNALSYM SHCNF_ACCEPT_INTERRUPTS}
  SHCNF_ACCEPT_INTERRUPTS     = $0001;
  {$EXTERNALSYM SHCNF_ACCEPT_NON_INTERRUPTS}
  SHCNF_ACCEPT_NON_INTERRUPTS = $0002;
  {$EXTERNALSYM SHCNF_NO_PROXY}
  SHCNF_NO_PROXY              = $8000;

type
  // Structures for the undocumented ChangeNotify handler
  TNotifyRegister = packed record // Structure that is passed to the SHChangeNotifyRegister Function
    ItemIDList: PItemIDList;
    bWatchSubTree: Bool;
  end;

  PDWordItemID = ^TDWordItemID; // Structure is what is passed in the wParam of the notify message when the notification is FreeSpace, ImageUpdate or anything with the SHCNF_DWORD flag.
  TDWordItemID = packed record
    cb: Word; { Size of Structure }
    dwItem1: DWORD;
    dwItem2: DWORD;
  end;

  PShellNotifyRec = ^TShellNotifyRec; // Structure is what is passed in the wParam of the notify message when the notification is anything with the SHCNF_IDLIST flag.
  TShellNotifyRec = packed record
    PIDL1,                           // Most ne_xxxx Notifications
    PIDL2: PItemIDList;
  end;

//------------------------------------------------------------------------------
// IContextMenu interfaces redefined to take advanatage of the Unicode Support
//------------------------------------------------------------------------------
type
  IContextMenu = interface(IUnknown)
    [SID_IContextMenu]
    function QueryContextMenu(Menu: HMENU;
      indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HResult; stdcall;
    function InvokeCommand(var lpici: TCMInvokeCommandInfoEx): HResult; stdcall;
    function GetCommandString(idCmd, uType: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HResult; stdcall;
  end;

  IContextMenu2 = interface(IContextMenu)
    [SID_IContextMenu2]
    function HandleMenuMsg(uMsg: UINT; WParam: WPARAM; LParam: LPARAM): HResult; stdcall;
  end;

  IContextMenu3 = interface(IContextMenu2)
  [SID_IContextMenu3]
    function HandleMenuMsg2(uMsg: UINT; wParam: WPARAM; lParam: LPARAM;
      var lpResult: LRESULT): HResult; stdcall;
  end;

type
  IShellIconOverlay = interface(IUnknown)
  [SID_IShellIconOverlay]
    function GetOverlayIndex(pidl: PItemIDList; var pIndex): HResult; stdcall;
    function GetOverlayIconIndex(pidl: PItemIDList; var pIconIndex): HResult; stdcall;
  end;

//------------------------------------------------------------------------------
// Button Constants
//------------------------------------------------------------------------------
const
  {$EXTERNALSYM MK_ALT}
  MK_ALT = $0020;
  {$EXTERNALSYM MK_BUTTON}
  MK_BUTTON = MK_LBUTTON or MK_RBUTTON or MK_MBUTTON;
  {$EXTERNALSYM MK_XBUTTON1}
  MK_XBUTTON1 = $0001;
  {$EXTERNALSYM MK_XBUTTON2}
  MK_XBUTTON2 = $0002;

//------------------------------------------------------------------------------
// Listview Column Constants
//------------------------------------------------------------------------------
const
  {$EXTERNALSYM LVCFMT_LEFT}
  LVCFMT_LEFT             = $0000;
  {$EXTERNALSYM LVCFMT_RIGHT}
  LVCFMT_RIGHT            = $0001;
  {$EXTERNALSYM LVCFMT_CENTER}
  LVCFMT_CENTER           = $0002;
  {$EXTERNALSYM LVCFMT_COL_HAS_IMAGES}
  LVCFMT_COL_HAS_IMAGES   = $8000;

//------------------------------------------------------------------------------
// New IShellFolder Constants and Types
//------------------------------------------------------------------------------
  {$EXTERNALSYM SFGAO_CANMONIKER}
  SFGAO_CANMONIKER = $400000;      // Defunct
  {$EXTERNALSYM SFGAO_HASSTORAGE}
  SFGAO_HASSTORAGE = $400000;      // Defunct
  {$EXTERNALSYM SFGAO_ENCRYPTED}
  SFGAO_ENCRYPTED = $2000;
  {$EXTERNALSYM SFGAO_ISSLOW}
  SFGAO_ISSLOW = $4000;
  {$EXTERNALSYM SFGAO_STORAGE}
  SFGAO_STORAGE = $0008;          // supports BindToObject(IID_IStorage)
  {$EXTERNALSYM SFGAO_STORAGEANCESTOR}
  SFGAO_STORAGEANCESTOR = $800000;// may contain children with SFGAO_STORAGE or SFGAO_STREAM
  {$EXTERNALSYM SFGAO_STREAM}
  SFGAO_STREAM = $400000;         // supports BindToObject(IID_IStream)



//------------------------------------------------------------------------------
// IShellIcon Constants and Types
//------------------------------------------------------------------------------
const
  // Constants for IShellIcon interface
  ICON_BLANK = 0;           // Unassocaiated file
  ICON_DATA  = 1;           // With data
  ICON_APP   = 2;           // Application, bat file etc
  ICON_FOLDER = 3;          // Plain folder
  ICON_FOLDEROPEN = 4;      // Open Folder

//------------------------------------------------------------------------------
// Drag Drop Image Helper Interfaces
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM IID_IDropTargetHelper}
  IID_IDropTargetHelper: TGUID = (
    D1:$4657278B; D2:$411B; D3:$11d2; D4:($83,$9A,$00,$C0,$4F,$D9,$18,$D0));
  {$EXTERNALSYM IID_IDragSourceHelper}
  IID_IDragSourceHelper: TGUID = (
    D1:$DE5BF786; D2:$477A; D3:$11d2; D4:($83,$9D,$00,$C0,$4F,$D9,$18,$D0));
  {$EXTERNALSYM CLSID_DragDropHelper}
  CLSID_DragDropHelper: TGUID = (
    D1:$4657278A; D2:$411B; D3:$11d2; D4:($83,$9A,$00,$C0,$4F,$D9,$18,$D0));

  {$EXTERNALSYM SID_IDropTargetHelper}
  SID_IDropTargetHelper = '{4657278B-411B-11d2-839A-00C04FD918D0}';
  {$EXTERNALSYM SID_IDragSourceHelper}
  SID_IDragSourceHelper = '{DE5BF786-477A-11d2-839D-00C04FD918D0}';
  {$EXTERNALSYM SID_IDropTarget}
  SID_IDropTarget = '{00000122-0000-0000-C000-000000000046}';

type
  IDropTargetHelper = interface(IUnknown)
    [SID_IDropTargetHelper]
    function DragEnter(hwndTarget: HWND; pDataObject: IDataObject; var ppt: TPoint; dwEffect: Integer): HRESULT; stdcall;
    function DragLeave: HRESULT; stdcall;
    function DragOver(var ppt: TPoint; dwEffect: Integer): HRESULT; stdcall;
    function Drop(pDataObject: IDataObject; var ppt: TPoint; dwEffect: Integer): HRESULT; stdcall;
    function Show(fShow: Boolean): HRESULT; stdcall;
  end;

  PSHDragImage = ^TSHDragImage;
  TSHDragImage = {$IFNDEF CPUX64}packed{$ENDIF} record
    sizeDragImage: TSize;
    ptOffset: TPoint;
    hbmpDragImage: HBITMAP;
    ColorRef: TColorRef;
  end;

  {$EXTERNALSYM IDragSourceHelper}
  IDragSourceHelper = interface(IUnknown)
    [SID_IDragSourceHelper]
    function InitializeFromBitmap(var SHDragImage: TSHDragImage; pDataObject: IDataObject): HRESULT; stdcall;
    function InitializeFromWindow(Window: HWND; var ppt: TPoint; pDataObject: IDataObject): HRESULT; stdcall;
  end;
  
//------------------------------------------------------------------------------
// IExtractImage definitions.
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM SID_IExtractImage}
  SID_IExtractImage = '{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}';
  {$EXTERNALSYM SID_IExtractImage2}
  SID_IExtractImage2 = '{953BB1EE-93B4-11d1-98A3-00C04FB687DA}';
  {$EXTERNALSYM IID_IExtractImage}
  IID_IExtractImage: TGUID = SID_IExtractImage;
  {$EXTERNALSYM IID_IExtractImage2}
  IID_IExtractImage2: TGUID = SID_IExtractImage2;

type
  IExtractImage = interface(IUnknown)
    [SID_IExtractImage]
    function GetLocation(Buffer: PWideChar;
                         BufferSize: DWORD;
                         var Priority: DWORD;
                         var Size: TSize;
                         ColorDepth: DWORD;
                         var Flags: DWORD ): HResult; stdcall;
    function Extract(var BmpImage: HBITMAP): HResult; stdcall;
  end;

  IExtractImage2 = interface(IExtractImage)
    [SID_IExtractImage2]
    function GetTimeStamp(var DateStamp: TFILETIME): HResult; stdcall;
  end;

const
  {$EXTERNALSYM IEI_PRIORITY_MAX}
  IEI_PRIORITY_MAX     = $0002;
  {$EXTERNALSYM IEI_PRIORITY_MIN}
  IEI_PRIORITY_MIN     = $0001;
  {$EXTERNALSYM IEI_PRIORITY_NORMAL}
  IEI_PRIORITY_NORMAL  = $0000;

  {$EXTERNALSYM IEIFLAG_ASYNC}
  IEIFLAG_ASYNC    = $0001;     // ask the extractor if it supports ASYNC extract (free threaded)
  {$EXTERNALSYM IEIFLAG_CACHE}
  IEIFLAG_CACHE    = $0002;     // returned from the extractor if it does NOT cache the thumbnail
  {$EXTERNALSYM IEIFLAG_ASPECT}
  IEIFLAG_ASPECT   = $0004;     // passed to the extractor to beg it to render to the aspect ratio of the supplied rect
  {$EXTERNALSYM IEIFLAG_OFFLINE}
  IEIFLAG_OFFLINE  = $0008;     // if the extractor shouldn't hit the net to get any content neede for the rendering
  {$EXTERNALSYM IEIFLAG_GLEAM}
  IEIFLAG_GLEAM    = $0010;     // does the image have a gleam ? this will be returned if it does
  {$EXTERNALSYM IEIFLAG_SCREEN}
  IEIFLAG_SCREEN   = $0020;     // render as if for the screen  (this is exlusive with IEIFLAG_ASPECT )
  {$EXTERNALSYM IEIFLAG_ORIGSIZE}
  IEIFLAG_ORIGSIZE = $0040;     // render to the approx size passed, but crop ifneccessary
  {$EXTERNALSYM IEIFLAG_NOSTAMP}
  IEIFLAG_NOSTAMP = $0080;     // returned from the extractor if it does NOT want an icon stamp on the thumbnail
  {$EXTERNALSYM IEIFLAG_NOBORDER}
  IEIFLAG_NOBORDER = $0100;    // returned from the extractor if it does NOT want an a border around the thumbnail
  {$EXTERNALSYM IEIFLAG_QUALITY}
  IEIFLAG_QUALITY = $0200;    // passed to the Extract method to indicate that a slower, higher quality image is desired, re-compute the thumbnail
  {$EXTERNALSYM IEIFLAG_REFRESH}
  IEIFLAG_REFRESH = $0400;   // returned from the extractor if it would like to have Refresh Thumbnail available



//------------------------------------------------------------------------------
// IShellLink definitions.
//------------------------------------------------------------------------------

const
 {$EXTERNALSYM CLSID_ShellLinkW}
  CLSID_ShellLinkW: TGUID = (
    D1:$000214F9; D2:$0000; D3:$0000; D4:($C0,$00,$00,$00,$00,$00,$00,$46));

  { IShellLink HotKey Mofifiers }
  {$EXTERNALSYM HOTKEYF_SHIFT}
  HOTKEYF_SHIFT =       $01;
  {$EXTERNALSYM HOTKEYF_CONTROL}
  HOTKEYF_CONTROL =     $02;
  {$EXTERNALSYM HOTKEYF_ALT}
  HOTKEYF_ALT =         $04;
  {$EXTERNALSYM HOTKEYF_EXT}
  HOTKEYF_EXT =         $08;


{$IFNDEF DELPHI_7_UP}
// D7 defines this right, D5 and D6 do not
// _FILEGROUPDESCRIPTORW Corrected definitions.
type
  PFileGroupDescriptorW = ^TFileGroupDescriptorW;
  {$EXTERNALSYM _FILEGROUPDESCRIPTORW}
  _FILEGROUPDESCRIPTORW = packed record
    cItems: UINT;
    fgd: array[0..0] of TFileDescriptorW;
  end;
  TFileGroupDescriptorW = _FILEGROUPDESCRIPTORW;
{$ENDIF}

// IShellLinkW Corrected definitions.
type
  {$EXTERNALSYM IShellLinkW}
  IShellLinkW = interface(IUnknown) { sl }
    [SID_IShellLinkW]
    function GetPath(pszFile: PWideChar; cchMaxPath: Integer;
      var pfd: TWin32FindDataW; fFlags: DWORD): HResult; stdcall;
    function GetIDList(var ppidl: PItemIDList): HResult; stdcall;
    function SetIDList(pidl: PItemIDList): HResult; stdcall;
    function GetDescription(pszName: PWideChar; cchMaxName: Integer): HResult; stdcall;
    function SetDescription(pszName: PWideChar): HResult; stdcall;
    function GetWorkingDirectory(pszDir: PWideChar; cchMaxPath: Integer): HResult; stdcall;
    function SetWorkingDirectory(pszDir: PWideChar): HResult; stdcall;
    function GetArguments(pszArgs: PWideChar; cchMaxPath: Integer): HResult; stdcall;
    function SetArguments(pszArgs: PWideChar): HResult; stdcall;
    function GetHotkey(var pwHotkey: Word): HResult; stdcall;
    function SetHotkey(wHotkey: Word): HResult; stdcall;
    function GetShowCmd(out piShowCmd: Integer): HResult; stdcall;
    function SetShowCmd(iShowCmd: Integer): HResult; stdcall;
    function GetIconLocation(pszIconPath: PWideChar; cchIconPath: Integer;
      out piIcon: Integer): HResult; stdcall;
    function SetIconLocation(pszIconPath: PWideChar; iIcon: Integer): HResult; stdcall;
    function SetRelativePath(pszPathRel: PWideChar; dwReserved: DWORD): HResult; stdcall;
    function Resolve(Wnd: HWND; fFlags: DWORD): HResult; stdcall;
    function SetPath(pszFile: PWideChar): HResult; stdcall;
  end;

//------------------------------------------------------------------------------
// IShellDetails definitions.
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM SID_IShellDetails}
  SID_IShellDetails = '{000214EC-0000-0000-C000-000000000046}';
  {$EXTERNALSYM IID_IShellDetails}
  IID_IShellDetails: TGUID = SID_IShellDetails;

type
  {$EXTERNALSYM PShellDetails}
  PShellDetails=^TShellDetails;
  {$EXTERNALSYM tagSHELLDETAILS}
  tagSHELLDETAILS = {$IFNDEF CPUX64}packed{$ENDIF} record
    Fmt: Integer;
    cxChar: Integer;
    str: TStrRet;
  end;
  {$EXTERNALSYM TShellDetails}
  TShellDetails = tagSHELLDETAILS;

  // BCB6 is all screwed up with IShellDetails. Can't get it to not see
  // my definition and it dissagrees that mine is right compared to the h file
  IShellDetailsBCB6=interface(IUnknown)
    [SID_IShellDetails]
    function GetDetailsOf(pidl: PItemIDList; iColumn: UINT; var pDetails: TShellDetails): HResult; stdcall;
    function ColumnClick(iColumn: LongWord): HResult; stdcall;
  end;

//------------------------------------------------------------------------------
// IShellFolder2 definitions.
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM IID_IEnumExtraSearch}
  IID_IEnumExtraSearch: TGUID = (
    D1:$0E700BE1; D2:$9DB6; D3:$11D1; D4:($A1,$CE,$00,$00,$4F,$D7,$5D,$13));
  {$EXTERNALSYM IID_IShellFolder2}
  IID_IShellFolder2: TGUID = (
    D1:$93F2F68C; D2:$1D1B; D3:$11D3; D4:($A3,$0E,$00,$C0,$4F,$79,$AB,$D1));
  IID_IPersistFolder3: TGUID = (
    D1:$CEF04FDF; D2:$FE72; D3:$11D2; D4:($87, $A5, $0, $C0, $4F, $68, $37, $CF));
  {$EXTERNALSYM SID_IPersistFolder3}
  SID_IPersistFolder3    = '{CEF04FDF-FE72-11D2-87A5-00C04F6837CF}';
  {$EXTERNALSYM IID_ITaskbarList}
  IID_ITaskbarList: TGUID = (
    D1:$56FDF342; D2:$FD6D; D3:$11D0; D4:($95,$8A,$00,$60,$97,$C9,$A0,$90));
  {$EXTERNALSYM IID_IDropTarget}
  IID_IDropTarget: TGUID = (
    D1:$00000122; D2:$0000; D3:$0000; D4:($C0,$00,$00,$00,$00,$00,$00,$46));

  CLSID_TaskbarList: TGUID = (
    D1:$56FDF344; D2:$FD6D; D3:$11D0; D4:($95,$8A,$00,$60,$97,$C9,$A0,$90));

  {$EXTERNALSYM SID_IShellFolder2}
  SID_IShellFolder2     = '{93F2F68C-1D1B-11D3-A30E-00C04F79ABD1}';
  {$EXTERNALSYM SID_IEnumExtraSearch}
  SID_IEnumExtraSearch  = '{0E700BE1-9DB6-11D1-A1CE-00004FD75D13}';
  {$EXTERNALSYM CLSID_TaskbarList}

const
  {$EXTERNALSYM SHCOLSTATE_TYPE_STR}
  SHCOLSTATE_TYPE_STR     = $00000001;
  {$EXTERNALSYM SHCOLSTATE_TYPE_INT}
  SHCOLSTATE_TYPE_INT     = $00000002;
  {$EXTERNALSYM SHCOLSTATE_TYPE_DATE}
  SHCOLSTATE_TYPE_DATE    = $00000003;
  {$EXTERNALSYM SHCOLSTATE_TYPEMASK}
  SHCOLSTATE_TYPEMASK     = $0000000F;
  {$EXTERNALSYM SHCOLSTATE_ONBYDEFAULT}
  SHCOLSTATE_ONBYDEFAULT  = $00000010;   // should on by default in details view
  {$EXTERNALSYM SHCOLSTATE_TYPE_SLOW}
  SHCOLSTATE_TYPE_SLOW    = $00000020;   // will be slow to compute, do on a background thread
 {$EXTERNALSYM SHCOLSTATE_EXTENDED}
  SHCOLSTATE_EXTENDED     = $00000040;   // provided by a handler, not the folder
 {$EXTERNALSYM SHCOLSTATE_SECONDARYUI}
  SHCOLSTATE_SECONDARYUI  = $00000080;   // not displayed in context menu, but listed in the "More..." dialog
 {$EXTERNALSYM SHCOLSTATE_HIDDEN}
  SHCOLSTATE_HIDDEN       = $00000100;   // not displayed in the UI

type
  tagEXTRASEARCH = packed record
    guidSearch: TGUID;
    wszFriendlyName: array[0..79] of WideChar;
    wszUrl: array[0..2083] of WideChar;
  end;
  PExtraSearch = ^TExtraSearch;
  TExtraSearch = tagEXTRASEARCH;

  {$EXTERNALSYM IEnumExtraSearch}
  IEnumExtraSearch = interface(IUnknown)
    [SID_IEnumExtraSearch]
    function Next(celt: ULONG; out rgElt: tagEXTRASEARCH; out pceltFetched: ULONG): HRESULT; stdcall;
    function Skip(celt: ULONG): HRESULT; stdcall;
    function Reset: HRESULT; stdcall;
    function Clone(out ppEnum: IEnumExtraSearch): HRESULT; stdcall;
  end;

  tagSHCOLUMNID = packed record
    fmtid: TGUID;
    pid: Cardinal;
  end;
  PSHColumnID = ^TSHColumnID;
  TSHColumnID = tagSHCOLUMNID;

  {$EXTERNALSYM IShellFolder2}
  IShellFolder2 = interface(IShellFolder)
    [SID_IShellFolder2]
    function GetDefaultSearchGUID(out pguid: TGUID): HRESULT; stdcall;
    function EnumSearches(out ppEnum: IEnumExtraSearch): HRESULT; stdcall;
    function GetDefaultColumn (dwRes: DWORD; out pSort: ULONG; out pDisplay: ULONG): HRESULT; stdcall;
    function GetDefaultColumnState(iColumn: UINT; out pcsFlags: DWORD): HRESULT; stdcall;
    function GetDetailsEx(pidl: PItemIDList; const pscid: TSHCOLUMNID; out pv: OLEVariant): HRESULT; stdcall;
    function GetDetailsOf(pidl: PItemIDList; iColumn: UINT; out psd: tagSHELLDETAILS): HRESULT; stdcall;
    function MapColumnToSCID(iColumn: UINT; out pscid: tagSHCOLUMNID): HRESULT; stdcall;
  end;

type
  PERSIST_FOLDER_TARGET_INFO = record
    pidlTargetFolder : PItemIdList;                           // pidl for the folder we want to intiailize
    szTargetParsingName : array [0..MAX_PATH-1] of WideChar;  // optional parsing name for the target
    szNetworkProvider : array [0..MAX_PATH-1] of WideChar;    // optional network provider
    dwAttributes : DWORD;                                     // optional FILE_ATTRIBUTES_ flags (-1 if not used)
    csidl : integer;                                          // optional folder index (SHGetFolderPath()) -1 if not used
  end;
  TPersistFolderTargetInfo = PERSIST_FOLDER_TARGET_INFO;
  PPersistFolderTargetInfo = ^PERSIST_FOLDER_TARGET_INFO;

  {$EXTERNALSYM IPersistFolder3}
  IPersistFolder3 = interface(IPersistFolder2)
    [SID_IPersistFolder3]
    function InitializeEx(pbc : IBindCtx; pidlRoot : PItemIdList; const ppfti : TPersistFolderTargetInfo) : HResult; stdcall;
    function GetFolderTargetInfo(var ppfti : TPersistFolderTargetInfo) : HResult; stdcall;
  end;

//------------------------------------------------------------------------------
// IColumnProvider types
//------------------------------------------------------------------------------

const
  SID_IColumnProvider = '{E8025004-1C42-11d2-BE2C-00A0C9A83DA1}';

  {$EXTERNALSYM IID_IColumnProvider}
  IID_IColumnProvider: TGUID = (
    D1:$E8025004; D2:$1C42; D3:$11D2; D4:($BE,$2C,$00,$A0,$C9,$A8,$3D,$A1));
  {$EXTERNALSYM CLSID_DocFileColumnProvider}
  CLSID_DocFileColumnProvider: TGUID = (
    D1:$24F14F01; D2:$7B1C; D3:$11D1; D4:($83,$8F,$00,$00,$F8,$04,$61,$CF));
  {$EXTERNALSYM CLSID_LinkColumnProvider}
  CLSID_LinkColumnProvider: TGUID = (
    D1:$24F14F02; D2:$7B1C; D3:$11D1; D4:($83,$8F,$00,$00,$F8,$04,$61,$CF));
  {$EXTERNALSYM CLSID_FileSysColumnProvider}
  CLSID_FileSysColumnProvider: TGUID = (
    D1:$0D2E74C4; D2:$3C34; D3:$11D2; D4:($A2,$7E,$00,$C0,$4F,$C3,$08,$71));

const

  {$EXTERNALSYM MAX_COLUMN_NAME_LEN}
  MAX_COLUMN_NAME_LEN = 80;    // Windows Defined
  {$EXTERNALSYM MAX_COLUMN_DESC_LEN}
  MAX_COLUMN_DESC_LEN = 128;   // Windows Defined

type
  PSHColumnInit = ^TSHColumnInit;
  TSHColumnInit = record
    dwFlags: ULONG;
    dwReserved: ULONG;
    wszFolder: packed array[0..MAX_PATH-1] of WideChar;
  end;

  TSHColumnInfo = packed record
    scid: TSHColumnID;     // Unique identifier for column
    vt: TVarType;          // Variant type
    fmt: DWORD;            // Alignment of the column, LVCFMT_xxx constants (ListviewColumn Format)
    cChars: UINT;          // Default Width of Column, in characters
    csFlags: DWORD;        // Default Column State
    wszTitle: array[0..MAX_COLUMN_NAME_LEN-1] of WideChar;  // Title of the Column
    wszDescription: array[0..MAX_COLUMN_DESC_LEN-1] of WideChar; // Description of the Column
  end;

const
  {$EXTERNALSYM SHCDF_UPDATEITEM}
  SHCDF_UPDATEITEM =   $00000001;    // this flag is a hint that the file has changed since the last call to GetItemData

  type
  PSHColumnData = ^TSHColumnData;
  TSHColumnData = record
    dwFlags: ULONG;
    dwFileAttributes: DWORD;
    dwReserved: ULONG;
    pwszExt: PWideChar;
    wszFile: packed array[0..MAX_PATH-1] of WideChar;
  end;

  {$EXTERNALSYM IColumnProvider}
  IColumnProvider = interface(IUnknown)
  [SID_IColumnProvider]
    function Initialize(psci: PSHColumnInit): HResult; stdcall;
    function GetColumnInfo(dwIndex: Longword; out psci: TSHColumnInfo): HResult; stdcall;
    function GetItemData(pscid: PSHColumnID; pscd: PSHColumnData; out pvarData: OLEVariant): HResult; stdcall;
  end;
//------------------------------------------------------------------------------
// SpecialFolder constants
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM CSIDL_COMMON_ADMINTOOLS}
  CSIDL_COMMON_ADMINTOOLS = $002f;     // All Users\Start Menu\Programs\Administrative Tools
  {$EXTERNALSYM CSIDL_ADMINTOOLS}
  CSIDL_ADMINTOOLS      = $0030;
  {$EXTERNALSYM CSIDL_LOCAL_APPDATA}
  CSIDL_LOCAL_APPDATA   = $001C;      // non roaming, user\Local Settings\Application Data
  {$EXTERNALSYM CSIDL_COOKIES}
  CSIDL_COOKIES         = $0021;
  {$EXTERNALSYM CSIDL_COMMON_APPDATA}
  CSIDL_COMMON_APPDATA  = $0023;      // All Users\Application Data
  {$EXTERNALSYM CSIDL_COMMON_TEMPLATES}
  CSIDL_COMMON_TEMPLATES = $002D;
  {$EXTERNALSYM CSIDL_WINDOWS}
  CSIDL_WINDOWS         = $0024;      // GetWindowsDirectory()
  {$EXTERNALSYM CSIDL_SYSTEM}
  CSIDL_SYSTEM          = $0025;      // GetSystemDirectory()
  {$EXTERNALSYM CSIDL_PROFILE}
  CSIDL_PROFILE         = $0028;      // USERPROFILE
  {$EXTERNALSYM CSIDL_PROGRAM_FILES}
  CSIDL_PROGRAM_FILES   = $0026;      // C:\Program Files
  {$EXTERNALSYM CSIDL_MYPICTURES}
  CSIDL_MYPICTURES      = $0027;      // My Pictures, new for Win2K
  {$EXTERNALSYM CSIDL_PROGRAM_FILES_COMMON}
  CSIDL_PROGRAM_FILES_COMMON = $002b; // C:\Program Files\Common
  {$EXTERNALSYM CSIDL_COMMON_DOCUMENTS}
  CSIDL_COMMON_DOCUMENTS = $002E;

//------------------------------------------------------------------------------
// AutoComplete definitions
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM CLSID_AutoComplete}
  CLSID_AutoComplete: TGUID = (
    D1:$00BB2763; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM CLSID_ACLHistory}
  CLSID_ACLHistory: TGUID = (
    D1:$00BB2764; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM CLSID_ACListISF}
  CLSID_ACListISF: TGUID = (
    D1:$03C036F1; D2:$A186; D3:$11D0; D4:($82,$4A,$00,$AA,$00,$5B,$43,$83));
  {$EXTERNALSYM CLSID_ACLMRU}
  CLSID_ACLMRU: TGUID = (
    D1:$6756a641; D2:$DE71; D3:$11D0; D4:($83,$1B,$00,$AA,$00,$5B,$43,$83));
  {$EXTERNALSYM CLSID_ACLMulti}
  CLSID_ACLMulti: TGUID = (
    D1:$00BB2765; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));

  {$EXTERNALSYM IID_IAutoComplete}
  IID_IAutoComplete: TGUID = (
    D1:$00BB2762; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM IID_IAutoComplete2}
  IID_IAutoComplete2: TGUID = (
    D1:$EAC04BC0; D2:$3791; D3:$11D2; D4:($BB,$95,$00,$60,$97,$7B,$46,$4C));
  {$EXTERNALSYM IID_IAutoCompList}
  IID_IAutoCompList: TGUID = (
    D1:$00BB2760; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM IID_IObjMgr}
  IID_IObjMgr: TGUID = (
    D1:$00BB2761; D2:$6A77; D3:$11D0; D4:($A5,$35,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM IID_IACList}
  IID_IACList: TGUID = (
    D1:$77A130B0; D2:$94FD; D3:$11D0; D4:($A5,$44,$00,$C0,$4F,$D7,$D0,$62));
  {$EXTERNALSYM IID_IACList2}
  IID_IACList2: TGUID = (
    D1:$470141A0; D2:$5186; D3:$11D2; D4:($BB,$B6,$00,$60,$97,$7B,$46,$4C));
  {$EXTERNALSYM IID_ICurrentWorkingDirectory}
  IID_ICurrentWorkingDirectory: TGUID = (
    D1:$91956d21; D2:$9276; D3:$11D1; D4:($92,$1A,$00,$60,$97,$DF,$5B,$D4));
  {$EXTERNALSYM IID_IEnumString}
  IID_IEnumString: TGUID = (
    D1:$00000101; D2:$0000; D3:$0000; D4:($C0,$00,$00,$00,$00,$00,$00,$46));

  {$EXTERNALSYM SID_IEmumString}
  SID_IEmumString              = '{00000101-0000-0000-C000-000000000046}';
  {$EXTERNALSYM SID_IAutoComplete}
  SID_IAutoComplete            = '{00BB2762-6A77-11D0-A535-00C04FD7D062}';
  {$EXTERNALSYM SID_IAutoComplete2}
  SID_IAutoComplete2           = '{EAC04BC0-3791-11D2-BB95-0060977B464C}';
  {$EXTERNALSYM SID_IACList}
  SID_IACList                  = '{77A130B0-94FD-11D0-A544-00C04FD7d062}';
  {$EXTERNALSYM SID_IACList2}
  SID_IACList2                 = '{470141A0-5186-11D2-BBB6-0060977B464C}';
  {$EXTERNALSYM SID_ICurrentWorkingDirectory}
  SID_ICurrentWorkingDirectory = '{91956D21-9276-11D1-921A-006097DF5BD4}';
  {$EXTERNALSYM SID_IObjMgr}
  SID_IObjMgr                  = '{00BB2761-6A77-11D0-A535-00C04FD7D062}';

  { AutoComplete2 Options }
  {$EXTERNALSYM ACO_NONE}
  ACO_NONE           = $0000;
  {$EXTERNALSYM ACO_AUTOSUGGEST}
  ACO_AUTOSUGGEST   = $0001;
  {$EXTERNALSYM ACO_AUTOAPPEND}
  ACO_AUTOAPPEND   = $0002;
  {$EXTERNALSYM ACO_SEARCH}
  ACO_SEARCH           = $0004;
  {$EXTERNALSYM ACO_FILTERPREFIXES}
  ACO_FILTERPREFIXES   = $0008;
  {$EXTERNALSYM ACO_USETAB}
  ACO_USETAB           = $0010;
  {$EXTERNALSYM ACO_UPDOWNKEYDROPSLIST}
  ACO_UPDOWNKEYDROPSLIST = $0020;
  {$EXTERNALSYM ACO_RTLREADING}
  ACO_RTLREADING   = $0040;

  // ACList2 Options
  {$EXTERNALSYM ACLO_NONE}
  ACLO_NONE            = $0000;    // don't enumerate anything
  {$EXTERNALSYM ACLO_CURRENTDIR}
  ACLO_CURRENTDIR      = $0001;    // enumerate current directory
  {$EXTERNALSYM ACLO_MYCOMPUTER}
  ACLO_MYCOMPUTER      = $0002;    // enumerate MyComputer
  {$EXTERNALSYM ACLO_DESKTOP}
  ACLO_DESKTOP         = $0004;    // enumerate Desktop Folder
  {$EXTERNALSYM ACLO_FAVORITES}
  ACLO_FAVORITES       = $0008;    // enumerate Favorites Folder
  {$EXTERNALSYM ACLO_FILESYSONLY}
  ACLO_FILESYSONLY     = $0010;   // enumerate only the file system
  {$EXTERNALSYM ACLO_FILESYSDIRS}
  ACLO_FILESYSDIRS     = $0020;   // Enumerate only the file system directories, Universal Naming Convention (UNC) shares, and UNC servers.


  ACLO_ALLOBJECTS = ACLO_CURRENTDIR or ACLO_MYCOMPUTER or ACLO_DESKTOP or ACLO_FAVORITES;

type
  IAutoComplete = interface(IUnknown)
    [SID_IAutoComplete]
    function Init( hWndEdit: HWND; punkACL: IUnknown; RegKeyPath, QuickComplete: POleStr): HRESULT; stdcall;
    function Enabled(fEnable: BOOL): HRESULT; stdcall;
  end;

  IAutoComplete2 = interface(IAutoComplete)
    ['{EAC04BC0-3791-11d2-BB95-0060977B464C}']
    function SetOptions(dwFlag: DWORD): HRESULT; stdcall;
    function GetOptions(out pdwFlag: DWORD): HRESULT; stdcall;
  end;

  IACList = interface(IUnknown)
    [SID_IACList]
    function Expand(pazExpand: LPCWSTR): HRESULT; stdcall;
  end;

  IACList2 = interface(IACList)
    [SID_IACList2]
    function SetOptions(pdwFlag: DWORD): HRESULT; stdcall;
    function GetOptions(var pdwFlag: DWORD): HRESULT; stdcall;
  end;

  ICurrentWorkingDirectory = interface(IUnknown)
    [SID_ICurrentWorkingDirectory]
    function GetDirectory(pwzPath: LPCWSTR; cchSize: DWORD): HRESULT; stdcall;
    function SetDirectory(pwzPath: LPCWSTR): HRESULT; stdcall;
  end;
//------------------------------------------------------------------------------

//------------------------------------------------------------------------------
// Win Base; Activation Context API
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID}
  ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID    = $00000001;
  {$EXTERNALSYM ACTCTX_FLAG_LANGID_VALID}
  ACTCTX_FLAG_LANGID_VALID                    = $00000002;
  {$EXTERNALSYM ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID}
  ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID        = $00000004;
  {$EXTERNALSYM ACTCTX_FLAG_RESOURCE_NAME_VALID}
  ACTCTX_FLAG_RESOURCE_NAME_VALID             = $00000008;
  {$EXTERNALSYM ACTCTX_FLAG_SET_PROCESS_DEFAULT}
  ACTCTX_FLAG_SET_PROCESS_DEFAULT             = $00000010;
  {$EXTERNALSYM ACTCTX_FLAG_APPLICATION_NAME_VALID}
  ACTCTX_FLAG_APPLICATION_NAME_VALID          = $00000020;
  {$EXTERNALSYM ACTCTX_FLAG_SOURCE_IS_ASSEMBLYREF}
  ACTCTX_FLAG_SOURCE_IS_ASSEMBLYREF           = $00000040;
  {$EXTERNALSYM ACTCTX_FLAG_HMODULE_VALID}
  ACTCTX_FLAG_HMODULE_VALID                   = $00000080;


  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_INTEL}
  PROCESSOR_ARCHITECTURE_INTEL           = 0;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_MIPS}
  PROCESSOR_ARCHITECTURE_MIPS            = 1;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_ALPHA}
  PROCESSOR_ARCHITECTURE_ALPHA           = 2;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_PPC}
  PROCESSOR_ARCHITECTURE_PPC             = 3;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_SHX}
  PROCESSOR_ARCHITECTURE_SHX             = 4;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_ARM}
  PROCESSOR_ARCHITECTURE_ARM             = 5;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_IA64}
  PROCESSOR_ARCHITECTURE_IA64            = 6;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_ALPHA64}
  PROCESSOR_ARCHITECTURE_ALPHA64         = 7;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_MSIL}
  PROCESSOR_ARCHITECTURE_MSIL            = 8;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_AMD64}
  PROCESSOR_ARCHITECTURE_AMD64           = 9;
  {$EXTERNALSYM PROCESSOR_ARCHITECTURE_IA32_ON_WIN64}
  PROCESSOR_ARCHITECTURE_IA32_ON_WIN64   = 10;


type
  tagACTCTXA = packed record
    cbSize: ULONG;
    dwFlags: DWORD;
    lpSource: LPCSTR;
    wProcessorArchitecture: WORD;
    wLangId: LANGID;
    lpAssemblyDirectory: LPCSTR;
    lpResourceName: LPCSTR;
    lpApplicationName: LPCSTR;
    hModule: HMODULE;
  end;
  TActCTXA = tagACTCTXA;
  PActCTXA = ^TActCTXA;

  tagACTCTXW = packed record
    cbSize: ULONG;
    dwFlags: DWORD;
    lpSource: LPCWSTR;
    wProcessorArchitecture: WORD;
    wLangId: LANGID;
    lpAssemblyDirectory: LPCWSTR;
    lpResourceName: LPCWSTR;
    lpApplicationName: LPCWSTR;
    hModule: HMODULE;
  end;
  TActCTXW = tagACTCTXW;
  PActCTXW = ^TActCTXW;

//------------------------------------------------------------------------------
// Property Page records with Activation Context fields
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM PSP_USEFUSIONCONTEXT}
  PSP_USEFUSIONCONTEXT      = $00004000;

type
  {$EXTERNALSYM _PROPSHEETPAGEA}
  _PROPSHEETPAGEA = record
    dwSize: Longint;
    dwFlags: Longint;
    hInstance: THandle;
    case Integer of
      0: (
        pszTemplate: PAnsiChar);
      1: (
        pResource: Pointer;
    case Integer of
      0: (
        hIcon: THandle);
      1: (
        pszIcon: PAnsiChar;
    pszTitle: PAnsiChar;
    pfnDlgProc: Pointer;
    lParam: LPARAM;
    pfnCallback: TFNPSPCallbackA;
    pcRefParent: PInteger;
    pszHeaderTitle: PAnsiChar;      // this is displayed in the header
    pszHeaderSubTitle: PAnsiChar;
    hActCtx: THandle)
  ); //

  end;
  {$EXTERNALSYM _PROPSHEETPAGEW}
  _PROPSHEETPAGEW = record
    dwSize: Longint;
    dwFlags: Longint;
    hInstance: THandle;
    case Integer of
      0: (
        pszTemplate: PWideChar);
      1: (
        pResource: Pointer;
    case Integer of
      0: (
        hIcon: THandle);
      1: (
        pszIcon: PWideChar;
    pszTitle: PWideChar;
    pfnDlgProc: Pointer;
    lParam: LPARAM;
    pfnCallback: TFNPSPCallbackW;
    pcRefParent: PInteger;
    pszHeaderTitle: PWideChar;      // this is displayed in the header
    pszHeaderSubTitle: PWideChar;
    hActCtx: THandle)); //
  end;
  {$EXTERNALSYM _PROPSHEETPAGE}
  _PROPSHEETPAGE = _PROPSHEETPAGEA;
  TPropSheetPageA = _PROPSHEETPAGEA;
  TPropSheetPageW = _PROPSHEETPAGEW;
  TPropSheetPage = TPropSheetPageA;
  {$EXTERNALSYM PROPSHEETPAGEA}
  PROPSHEETPAGEA = _PROPSHEETPAGEA;
  {$EXTERNALSYM PROPSHEETPAGEW}
  PROPSHEETPAGEW = _PROPSHEETPAGEW;
  {$EXTERNALSYM PROPSHEETPAGE}
  PROPSHEETPAGE = PROPSHEETPAGEA;

  TThreadProc = function(lpParameter: Pointer): DWORD; stdcall;

//------------------------------------------------------------------------------
// WebBrowser Interfaces
//------------------------------------------------------------------------------

const
  DIID_DWebBrowserEvents: TGUID = '{EAB22AC2-30C1-11CF-A7EB-0000C05BAE0B}';
  DIID_DWebBrowserEvents2: TGUID = '{34A715A0-6587-11D0-924A-0020AFC7AC4D}';
  CLSID_ToolbarExtButtons: TGUID = '{2CE4B5D8-A28F-11D2-86C5-00C04F8EEA99}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// CommandStateChangeConstants constants
type
  {$IFNDEF COMPILER_5_UP}
  TOleEnum = type Integer;
  {$ENDIF}
  
  CommandStateChangeConstants = TOleEnum;
const
  {$EXTERNALSYM CSC_UPDATECOMMANDS}
  CSC_UPDATECOMMANDS = $FFFFFFFF;
  {$EXTERNALSYM CSC_NAVIGATEFORWARD}
  CSC_NAVIGATEFORWARD = $00000001;
  {$EXTERNALSYM CSC_NAVIGATEBACK}
  CSC_NAVIGATEBACK = $00000002;

// OLECMDID constants
type
  OLECMDID = TOleEnum;
const
  {$EXTERNALSYM OLECMDID_OPEN}
  OLECMDID_OPEN = $00000001;
  {$EXTERNALSYM OLECMDID_NEW}
  OLECMDID_NEW = $00000002;
  {$EXTERNALSYM OLECMDID_SAVE}
  OLECMDID_SAVE = $00000003;
  {$EXTERNALSYM OLECMDID_SAVEAS}
  OLECMDID_SAVEAS = $00000004;
  {$EXTERNALSYM OLECMDID_SAVECOPYAS}
  OLECMDID_SAVECOPYAS = $00000005;
  {$EXTERNALSYM OLECMDID_PRINT}
  OLECMDID_PRINT = $00000006;
  {$EXTERNALSYM OLECMDID_PRINTPREVIEW}
  OLECMDID_PRINTPREVIEW = $00000007;
  {$EXTERNALSYM OLECMDID_PAGESETUP}
  OLECMDID_PAGESETUP = $00000008;
  {$EXTERNALSYM OLECMDID_SPELL}
  OLECMDID_SPELL = $00000009;
  {$EXTERNALSYM OLECMDID_PROPERTIES}
  OLECMDID_PROPERTIES = $0000000A;
  {$EXTERNALSYM OLECMDID_CUT}
  OLECMDID_CUT = $0000000B;
  {$EXTERNALSYM OLECMDID_COPY}
  OLECMDID_COPY = $0000000C;
  {$EXTERNALSYM OLECMDID_PASTE}
  OLECMDID_PASTE = $0000000D;
  {$EXTERNALSYM OLECMDID_PASTESPECIAL}
  OLECMDID_PASTESPECIAL = $0000000E;
  {$EXTERNALSYM OLECMDID_UNDO}
  OLECMDID_UNDO = $0000000F;
  {$EXTERNALSYM OLECMDID_REDO}
  OLECMDID_REDO = $00000010;
  {$EXTERNALSYM OLECMDID_SELECTALL}
  OLECMDID_SELECTALL = $00000011;
  {$EXTERNALSYM OLECMDID_CLEARSELECTION}
  OLECMDID_CLEARSELECTION = $00000012;
  {$EXTERNALSYM OLECMDID_ZOOM}
  OLECMDID_ZOOM = $00000013;
  {$EXTERNALSYM OLECMDID_GETZOOMRANGE}
  OLECMDID_GETZOOMRANGE = $00000014;
  {$EXTERNALSYM OLECMDID_UPDATECOMMANDS}
  OLECMDID_UPDATECOMMANDS = $00000015;
  {$EXTERNALSYM OLECMDID_REFRESH}
  OLECMDID_REFRESH = $00000016;
  {$EXTERNALSYM OLECMDID_STOP}
  OLECMDID_STOP = $00000017;
  {$EXTERNALSYM OLECMDID_HIDETOOLBARS}
  OLECMDID_HIDETOOLBARS = $00000018;
  {$EXTERNALSYM OLECMDID_SETPROGRESSMAX}
  OLECMDID_SETPROGRESSMAX = $00000019;
  {$EXTERNALSYM OLECMDID_SETPROGRESSPOS}
  OLECMDID_SETPROGRESSPOS = $0000001A;
  {$EXTERNALSYM OLECMDID_SETPROGRESSTEXT}
  OLECMDID_SETPROGRESSTEXT = $0000001B;
  {$EXTERNALSYM OLECMDID_SETTITLE}
  OLECMDID_SETTITLE = $0000001C;
  {$EXTERNALSYM OLECMDID_SETDOWNLOADSTATE}
  OLECMDID_SETDOWNLOADSTATE = $0000001D;
  {$EXTERNALSYM OLECMDID_STOPDOWNLOAD}
  OLECMDID_STOPDOWNLOAD = $0000001E;
  {$EXTERNALSYM OLECMDID_ONTOOLBARACTIVATED}
  OLECMDID_ONTOOLBARACTIVATED = $0000001F;
  {$EXTERNALSYM OLECMDID_FIND}
  OLECMDID_FIND = $00000020;
  {$EXTERNALSYM OLECMDID_DELETE}
  OLECMDID_DELETE = $00000021;
  {$EXTERNALSYM OLECMDID_HTTPEQUIV}
  OLECMDID_HTTPEQUIV = $00000022;
  {$EXTERNALSYM OLECMDID_HTTPEQUIV_DONE}
  OLECMDID_HTTPEQUIV_DONE = $00000023;
  {$EXTERNALSYM OLECMDID_ENABLE_INTERACTION}
  OLECMDID_ENABLE_INTERACTION = $00000024;
  {$EXTERNALSYM OLECMDID_ONUNLOAD}
  OLECMDID_ONUNLOAD = $00000025;
  {$EXTERNALSYM OLECMDID_PROPERTYBAG2}
  OLECMDID_PROPERTYBAG2 = $00000026;
  {$EXTERNALSYM OLECMDID_PREREFRESH}
  OLECMDID_PREREFRESH = $00000027;
  {$EXTERNALSYM OLECMDID_SHOWSCRIPTERROR}
  OLECMDID_SHOWSCRIPTERROR = $00000028;
  {$EXTERNALSYM OLECMDID_SHOWMESSAGE}
  OLECMDID_SHOWMESSAGE = $00000029;
  {$EXTERNALSYM OLECMDID_SHOWFIND}
  OLECMDID_SHOWFIND = $0000002A;
  {$EXTERNALSYM OLECMDID_SHOWPAGESETUP}
  OLECMDID_SHOWPAGESETUP = $0000002B;
  {$EXTERNALSYM OLECMDID_SHOWPRINT}
  OLECMDID_SHOWPRINT = $0000002C;
  {$EXTERNALSYM OLECMDID_CLOSE}
  OLECMDID_CLOSE = $0000002D;
  {$EXTERNALSYM OLECMDID_ALLOWUILESSSAVEAS}
  OLECMDID_ALLOWUILESSSAVEAS = $0000002E;
  {$EXTERNALSYM OLECMDID_DONTDOWNLOADCSS}
  OLECMDID_DONTDOWNLOADCSS = $0000002F;

// OLECMDF constants
type
  OLECMDF = TOleEnum;
const
  {$EXTERNALSYM OLECMDF_SUPPORTED}
  OLECMDF_SUPPORTED = $00000001;
  {$EXTERNALSYM OLECMDF_ENABLED}
  OLECMDF_ENABLED = $00000002;
  {$EXTERNALSYM OLECMDF_LATCHED}
  OLECMDF_LATCHED = $00000004;
  {$EXTERNALSYM OLECMDF_NINCHED}
  OLECMDF_NINCHED = $00000008;
  {$EXTERNALSYM OLECMDF_INVISIBLE}
  OLECMDF_INVISIBLE = $00000010;
  {$EXTERNALSYM OLECMDF_DEFHIDEONCTXTMENU}
  OLECMDF_DEFHIDEONCTXTMENU = $00000020;

// OLECMDEXECOPT constants
type
  OLECMDEXECOPT = TOleEnum;
const
  {$EXTERNALSYM OLECMDEXECOPT_DODEFAULT}
  OLECMDEXECOPT_DODEFAULT = $00000000;
  {$EXTERNALSYM OLECMDEXECOPT_PROMPTUSER}
  OLECMDEXECOPT_PROMPTUSER = $00000001;
  {$EXTERNALSYM OLECMDEXECOPT_DONTPROMPTUSER}
  OLECMDEXECOPT_DONTPROMPTUSER = $00000002;
  {$EXTERNALSYM OLECMDEXECOPT_SHOWHELP}
  OLECMDEXECOPT_SHOWHELP = $00000003;

// tagREADYSTATE constants
type
  tagREADYSTATE = TOleEnum;
const
  {$EXTERNALSYM READYSTATE_UNINITIALIZED}
  READYSTATE_UNINITIALIZED = $00000000;
  {$EXTERNALSYM READYSTATE_LOADING}
  READYSTATE_LOADING = $00000001;
  {$EXTERNALSYM READYSTATE_LOADED}
  READYSTATE_LOADED = $00000002;
  {$EXTERNALSYM READYSTATE_INTERACTIVE}
  READYSTATE_INTERACTIVE = $00000003;
  {$EXTERNALSYM READYSTATE_COMPLETE}
  READYSTATE_COMPLETE = $00000004;

// ShellWindowTypeConstants constants
type
  ShellWindowTypeConstants = TOleEnum;
const
  {$EXTERNALSYM SWC_EXPLORER}
  SWC_EXPLORER = $00000000;
  {$EXTERNALSYM SWC_BROWSER}
  SWC_BROWSER = $00000001;
  {$EXTERNALSYM SWC_3RDPARTY}
  SWC_3RDPARTY = $00000002;
  {$EXTERNALSYM SWC_CALLBACK}
  SWC_CALLBACK = $00000004;

// ShellWindowFindWindowOptions constants
type
  ShellWindowFindWindowOptions = TOleEnum;
const
  {$EXTERNALSYM SWFO_NEEDDISPATCH}
  SWFO_NEEDDISPATCH = $00000001;
  {$EXTERNALSYM SWFO_INCLUDEPENDING}
  SWFO_INCLUDEPENDING = $00000002;
  {$EXTERNALSYM SWFO_COOKIEPASSED}
  SWFO_COOKIEPASSED = $00000004;

type
  RefreshConstants = TOleEnum;
const
  {$EXTERNALSYM REFRESH_NORMAL}
  REFRESH_NORMAL = 0;
  {$EXTERNALSYM REFRESH_IFEXPIRED}
  REFRESH_IFEXPIRED = 1;
  {$EXTERNALSYM REFRESH_CONTINUE}
  REFRESH_CONTINUE = 2;
  {$EXTERNALSYM REFRESH_COMPLETELY}
  REFRESH_COMPLETELY = 3;

type
  BrowserNavConstants = TOleEnum;
const
  {$EXTERNALSYM navOpenInNewWindow}
  navOpenInNewWindow = $00000001;
  {$EXTERNALSYM navNoHistory}
  navNoHistory       = $00000002;
  {$EXTERNALSYM navNoReadFromCache}
  navNoReadFromCache = $00000004;
  {$EXTERNALSYM navNoWriteToCache}
  navNoWriteToCache  = $00000008;
  {$EXTERNALSYM navAllowAutosearch}
  navAllowAutosearch = $00000010;
  {$EXTERNALSYM navBrowserBar}
  navBrowserBar      = $00000020;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IWebBrowser = interface;
  IWebBrowser2 = interface;

// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}

// *********************************************************************//
// Interface: IWebBrowser
// *********************************************************************//
  IWebBrowser = interface(IDispatch)
    ['{EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B}']
    procedure GoBack; safecall;
    procedure GoForward; safecall;
    procedure GoHome; safecall;
    procedure GoSearch; safecall;
    procedure Navigate(const URL: WideString; var Flags: OleVariant;
                       var TargetFrameName: OleVariant; var PostData: OleVariant; 
                       var Headers: OleVariant); safecall;
    procedure Refresh; safecall;
    procedure Refresh2(var Level: OleVariant); safecall;
    procedure Stop; safecall;
    function  Get_Application: IDispatch; safecall;
    function  Get_Parent: IDispatch; safecall;
    function  Get_Container: IDispatch; safecall;
    function  Get_Document: IDispatch; safecall;
    function  Get_TopLevelContainer: WordBool; safecall;
    function  Get_Type_: WideString; safecall;
    function  Get_Left: Integer; safecall;
    procedure Set_Left(pl: Integer); safecall;
    function  Get_Top: Integer; safecall;
    procedure Set_Top(pl: Integer); safecall;
    function  Get_Width: Integer; safecall;
    procedure Set_Width(pl: Integer); safecall;
    function  Get_Height: Integer; safecall;
    procedure Set_Height(pl: Integer); safecall;
    function  Get_LocationName: WideString; safecall;
    function  Get_LocationURL: WideString; safecall;
    function  Get_Busy: WordBool; safecall;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
    property Container: IDispatch read Get_Container;
    property Document: IDispatch read Get_Document;
    property TopLevelContainer: WordBool read Get_TopLevelContainer;
    property Type_: WideString read Get_Type_;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property LocationName: WideString read Get_LocationName;
    property LocationURL: WideString read Get_LocationURL;
    property Busy: WordBool read Get_Busy;
  end;

// *********************************************************************//
// Interface: IWebBrowserApp
// *********************************************************************//
  IWebBrowserApp = interface(IWebBrowser)
    ['{0002DF05-0000-0000-C000-000000000046}']
    procedure Quit; safecall;
    procedure ClientToWindow(var pcx: SYSINT; var pcy: SYSINT); safecall;
    procedure PutProperty(const Property_: WideString; vtValue: OleVariant); safecall;
    function  GetProperty(const Property_: WideString): OleVariant; safecall;
    function  Get_Name: WideString; safecall;
    function  Get_HWND: Integer; safecall;
    function  Get_FullName: WideString; safecall;
    function  Get_Path: WideString; safecall;
    function  Get_Visible: WordBool; safecall;
    procedure Set_Visible(pBool: WordBool); safecall;
    function  Get_StatusBar: WordBool; safecall;
    procedure Set_StatusBar(pBool: WordBool); safecall;
    function  Get_StatusText: WideString; safecall;
    procedure Set_StatusText(const StatusText: WideString); safecall;
    function  Get_ToolBar: SYSINT; safecall;
    procedure Set_ToolBar(Value: SYSINT); safecall;
    function  Get_MenuBar: WordBool; safecall;
    procedure Set_MenuBar(Value: WordBool); safecall;
    function  Get_FullScreen: WordBool; safecall;
    procedure Set_FullScreen(pbFullScreen: WordBool); safecall;
    property Name: WideString read Get_Name;
    property HWND: Integer read Get_HWND;
    property FullName: WideString read Get_FullName;
    property Path: WideString read Get_Path;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property StatusBar: WordBool read Get_StatusBar write Set_StatusBar;
    property StatusText: WideString read Get_StatusText write Set_StatusText;
    property ToolBar: SYSINT read Get_ToolBar write Set_ToolBar;
    property MenuBar: WordBool read Get_MenuBar write Set_MenuBar;
    property FullScreen: WordBool read Get_FullScreen write Set_FullScreen;
  end;

// *********************************************************************//
// Interface: IWebBrowser2
// *********************************************************************//
  IWebBrowser2 = interface(IWebBrowserApp)
    ['{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}']
    procedure Navigate2(var URL: OleVariant; var Flags: OleVariant;
                        var TargetFrameName: OleVariant; var PostData: OleVariant;
                        var Headers: OleVariant); safecall;
    function  QueryStatusWB(cmdID: OLECMDID): OLECMDF; safecall;
    procedure ExecWB(cmdID: OLECMDID; cmdexecopt: OLECMDEXECOPT; var pvaIn: OleVariant;
                     var pvaOut: OleVariant); safecall;
    procedure ShowBrowserBar(var pvaClsid: OleVariant; var pvarShow: OleVariant;
                             var pvarSize: OleVariant); safecall;
    function  Get_ReadyState: tagREADYSTATE; safecall;
    function  Get_Offline: WordBool; safecall;
    procedure Set_Offline(pbOffline: WordBool); safecall;
    function  Get_Silent: WordBool; safecall;
    procedure Set_Silent(pbSilent: WordBool); safecall;
    function  Get_RegisterAsBrowser: WordBool; safecall;
    procedure Set_RegisterAsBrowser(pbRegister: WordBool); safecall;
    function  Get_RegisterAsDropTarget: WordBool; safecall;
    procedure Set_RegisterAsDropTarget(pbRegister: WordBool); safecall;
    function  Get_TheaterMode: WordBool; safecall;
    procedure Set_TheaterMode(pbRegister: WordBool); safecall;
    function  Get_AddressBar: WordBool; safecall;
    procedure Set_AddressBar(Value: WordBool); safecall;
    function  Get_Resizable: WordBool; safecall;
    procedure Set_Resizable(Value: WordBool); safecall;
    property ReadyState: tagREADYSTATE read Get_ReadyState;
    property Offline: WordBool read Get_Offline write Set_Offline;
    property Silent: WordBool read Get_Silent write Set_Silent;
    property RegisterAsBrowser: WordBool read Get_RegisterAsBrowser write Set_RegisterAsBrowser;
    property RegisterAsDropTarget: WordBool read Get_RegisterAsDropTarget write Set_RegisterAsDropTarget;
    property TheaterMode: WordBool read Get_TheaterMode write Set_TheaterMode;
    property AddressBar: WordBool read Get_AddressBar write Set_AddressBar;
    property Resizable: WordBool read Get_Resizable write Set_Resizable;
  end;

const
  DIID_DShellFolderViewEvents: TGUID = '{62112AA2-EBE4-11CF-A5FB-0020AFE7292D}';
  IID_FolderItemVerbs: TGUID = '{1F8352C0-50B0-11CF-960C-0080C7F4EE85}';
  IID_FolderItemVerb: TGUID = '{08EC3E00-50B0-11CF-960C-0080C7F4EE85}';
  IID_FolderItems: TGUID = '{744129E0-CBE5-11CE-8350-444553540000}';
  IID_FolderItems2: TGUID = '{C94F0AD0-F363-11D2-A327-00C04F8EEC7F}';
  IID_FolderItems3: TGUID = '{EAA7C309-BBEC-49D5-821D-64D966CB667F}';
  IID_Folder: TGUID = '{BBCBDE60-C3FF-11CE-8350-444553540000}';
  IID_Folder2: TGUID = '{F0D2D8EF-3890-11D2-BF8B-00C04FB93661}';
  IID_Folder3: TGUID = '{A7AE5F64-C4D7-4D7F-9307-4D24EE54B841}';

type
// *********************************************************************//
// Interface: FolderItemVerb
// GUID:      {08EC3E00-50B0-11CF-960C-0080C7F4EE85}
// *********************************************************************//
  FolderItemVerb = interface(IDispatch)
    ['{08EC3E00-50B0-11CF-960C-0080C7F4EE85}']
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    procedure DoIt; safecall;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// Interface: FolderItemVerbs
// GUID:      {1F8352C0-50B0-11CF-960C-0080C7F4EE85}
// *********************************************************************//
  FolderItemVerbs = interface(IDispatch)
    ['{1F8352C0-50B0-11CF-960C-0080C7F4EE85}']
    function Get_Count: Integer; safecall;
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(index: OleVariant): FolderItemVerb; safecall;
    function _NewEnum: IUnknown; safecall;
    property Count: Integer read Get_Count;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// Interface: FolderItem
// *********************************************************************//
  FolderItem = interface(IDispatch)
    ['{FAC32C80-CBE4-11CE-8350-444553540000}']
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pbs: WideString); safecall;
    function Get_Path: WideString; safecall;
    function Get_GetLink: IDispatch; safecall;
    function Get_GetFolder: IDispatch; safecall;
    function Get_IsLink: WordBool; safecall;
    function Get_IsFolder: WordBool; safecall;
    function Get_IsFileSystem: WordBool; safecall;
    function Get_IsBrowsable: WordBool; safecall;
    function Get_ModifyDate: TDateTime; safecall;
    procedure Set_ModifyDate(pdt: TDateTime); safecall;
    function Get_Size: Integer; safecall;
    function Get_type_: WideString; safecall;
    function Verbs: FolderItemVerbs; safecall;
    procedure InvokeVerb(vVerb: OleVariant); safecall;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name write Set_Name;
    property Path: WideString read Get_Path;
    property GetLink: IDispatch read Get_GetLink;
    property GetFolder: IDispatch read Get_GetFolder;
    property IsLink: WordBool read Get_IsLink;
    property IsFolder: WordBool read Get_IsFolder;
    property IsFileSystem: WordBool read Get_IsFileSystem;
    property IsBrowsable: WordBool read Get_IsBrowsable;
    property ModifyDate: TDateTime read Get_ModifyDate write Set_ModifyDate;
    property Size: Integer read Get_Size;
    property type_: WideString read Get_type_;
  end;

// *********************************************************************//
// Interface: FolderItem2
// *********************************************************************//
  FolderItem2 = interface(FolderItem)
    ['{EDC817AA-92B8-11D1-B075-00C04FC33AA5}']
    procedure InvokeVerbEx(vVerb: OleVariant; vArgs: OleVariant); safecall;
    function ExtendedProperty(const bstrPropName: WideString): OleVariant; safecall;
  end;

// *********************************************************************//
// Interface: FolderItems
// *********************************************************************//
  FolderItems = interface(IDispatch)
    ['{744129E0-CBE5-11CE-8350-444553540000}']
    function Get_Count: Integer; safecall;
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(index: OleVariant): FolderItem; safecall;
    function _NewEnum: IUnknown; safecall;
    property Count: Integer read Get_Count;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// Interface: FolderItems2
// *********************************************************************//
  FolderItems2 = interface(FolderItems)
    ['{C94F0AD0-F363-11D2-A327-00C04F8EEC7F}']
    procedure InvokeVerbEx(vVerb: OleVariant; vArgs: OleVariant); safecall;
  end;


// *********************************************************************//
// Interface: FolderItems3
// *********************************************************************//
  FolderItems3 = interface(FolderItems2)
    ['{EAA7C309-BBEC-49D5-821D-64D966CB667F}']
    procedure Filter(grfFlags: Integer; const bstrFileSpec: WideString); safecall;
    function Get_Verbs: FolderItemVerbs; safecall;
    property Verbs: FolderItemVerbs read Get_Verbs;
  end;


// *********************************************************************//
// Interface: Folder
// *********************************************************************//
  Folder = interface(IDispatch)
    ['{BBCBDE60-C3FF-11CE-8350-444553540000}']
    function Get_Title: WideString; safecall;
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ParentFolder: Folder; safecall;
    function Items: FolderItems; safecall;
    function ParseName(const bName: WideString): FolderItem; safecall;
    procedure NewFolder(const bName: WideString; vOptions: OleVariant); safecall;
    procedure MoveHere(vItem: OleVariant; vOptions: OleVariant); safecall;
    procedure CopyHere(vItem: OleVariant; vOptions: OleVariant); safecall;
    function GetDetailsOf(vItem: OleVariant; iColumn: SYSINT): WideString; safecall;
    property Title: WideString read Get_Title;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
    property ParentFolder: Folder read Get_ParentFolder;
  end;

// *********************************************************************//
// Interface: Folder2
// *********************************************************************//
  Folder2 = interface(Folder)
    ['{F0D2D8EF-3890-11D2-BF8B-00C04FB93661}']
    function Get_Self: FolderItem; safecall;
    function Get_OfflineStatus: Integer; safecall;
    procedure Synchronize; safecall;
    function Get_HaveToShowWebViewBarricade: WordBool; safecall;
    procedure DismissedWebViewBarricade; safecall;
    property Self: FolderItem read Get_Self;
    property OfflineStatus: Integer read Get_OfflineStatus;
    property HaveToShowWebViewBarricade: WordBool read Get_HaveToShowWebViewBarricade;
  end;

// *********************************************************************//
// Interface: Folder3
// *********************************************************************//
  Folder3 = interface(Folder2)
    ['{A7AE5F64-C4D7-4D7F-9307-4D24EE54B841}']
    function Get_ShowWebViewBarricade: WordBool; safecall;
    procedure Set_ShowWebViewBarricade(pbShowWebViewBarricade: WordBool); safecall;
    property ShowWebViewBarricade: WordBool read Get_ShowWebViewBarricade write Set_ShowWebViewBarricade;
  end;

// *********************************************************************//
// Interface: IShellFolderViewDual
// *********************************************************************//
  IShellFolderViewDual = interface(IDispatch)
    ['{E7A1AF80-4D96-11CF-960C-0080C7F4EE85}']
    function Get_Application: IDispatch; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Folder: Folder; safecall;
    function SelectedItems: FolderItems; safecall;
    function Get_FocusedItem: FolderItem; safecall;
    procedure SelectItem(var pvfi: OleVariant; dwFlags: SYSINT); safecall;
    function PopupItemMenu(const pfi: FolderItem; vx: OleVariant; vy: OleVariant): WideString; safecall;
    function Get_Script: IDispatch; safecall;
    function Get_ViewOptions: Integer; safecall;
    property Application: IDispatch read Get_Application;
    property Parent: IDispatch read Get_Parent;
    property Folder: Folder read Get_Folder;
    property FocusedItem: FolderItem read Get_FocusedItem;
    property Script: IDispatch read Get_Script;
    property ViewOptions: Integer read Get_ViewOptions;
  end;

{$IFDEF CPPB_6}
// *********************************************************************//
// DispIntf:  DShellFolderViewEvents
// *********************************************************************//
  DShellFolderViewEvents = dispinterface
    ['{62112AA2-EBE4-11CF-A5FB-0020AFE7292D}']
    procedure SelectionChanged; dispid 200;
    procedure EnumDone; dispid 201;
    function VerbInvoked: WordBool; dispid 202;
    function DefaultVerbInvoked: WordBool; dispid 203;
    function BeginDrag: WordBool; dispid 204;
  end;
{$ENDIF CPPB_6}

//------------------------------------------------------------------------------
// SHCreateShellFolderViewEx
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM FWF_SHOWSELALWAYS}
  FWF_SHOWSELALWAYS       = $2000;
  {$EXTERNALSYM FWF_NOVISIBLE}
  FWF_NOVISIBLE           = $4000;
  {$EXTERNALSYM FWF_SINGLECLICKACTIVATE}
  FWF_SINGLECLICKACTIVATE  = $8000;
  {$EXTERNALSYM FWF_NOWEBVIEW}
  FWF_NOWEBVIEW           = $10000;
  {$EXTERNALSYM FWF_HIDEFILENAMES}
  FWF_HIDEFILENAMES        = $20000;
  {$EXTERNALSYM FWF_CHECKSELECT}
  FWF_CHECKSELECT          = $40000;


  {$EXTERNALSYM SVUIA_DEACTIVATE}
  SVUIA_DEACTIVATE        = $0000;
  {$EXTERNALSYM SVUIA_ACTIVATE_NOFOCUS}
  SVUIA_ACTIVATE_NOFOCUS  = $0001;
  {$EXTERNALSYM SVUIA_ACTIVATE_FOCUS}
  SVUIA_ACTIVATE_FOCUS    = $0002;
  {$EXTERNALSYM SVUIA_INPLACEACTIVATE}
  SVUIA_INPLACEACTIVATE   = $0003;

  {$EXTERNALSYM FVM_ICON}
  FVM_ICON      = 1;
  {$EXTERNALSYM FVM_SMALLICON}
  FVM_SMALLICON  = 2;
  {$EXTERNALSYM FVM_LIST}
  FVM_LIST      = 3;
  {$EXTERNALSYM FVM_DETAILS}
  FVM_DETAILS    = 4;
  {$EXTERNALSYM FVM_THUMBNAIL}
  FVM_THUMBNAIL  = 5;
  {$EXTERNALSYM FVM_TILE}
  FVM_TILE      = 6;
  {$EXTERNALSYM FVM_THUMBSTRIP}
  FVM_THUMBSTRIP= 7;
  {$EXTERNALSYM FVM_LAST}
  FVM_LAST      = 7;


// For the SFVM_MERGEMENU message
const
  {$EXTERNALSYM QCMINFO_PLACE_BEFORE}
  QCMINFO_PLACE_BEFORE   = 0;
  {$EXTERNALSYM QCMINFO_PLACE_AFTER}
  QCMINFO_PLACE_AFTER    = 1;

type
  QCMINFO_IDMAP_PLACEMENT = record
    id : UINT;
    fFlags : UINT;
  end;
  PQCMINFO_IDMAP_PLACEMENT = ^QCMINFO_IDMAP_PLACEMENT;
  TQCMInfoIDMapPlacement = QCMINFO_IDMAP_PLACEMENT;
  PQCMInfoIDMapPlacement = PQCMINFO_IDMAP_PLACEMENT;

  QCMINFO_IDMAP = record
    nMaxIDs : UINT;
    pIdList : packed array [0..0] of QCMINFO_IDMAP_PLACEMENT;
  end;

  PQCMINFO_IDMAP = ^QCMINFO_IDMAP;
   QCMINFO = record
    menu : HMENU;             // in
    indexMenu : UINT;         // in
    idCmdFirst : UINT;        // in/out
    idCmdLast : UINT;         // in
    pIDMap : PQCMINFO_IDMAP;  // in / unused
  end;
  PQCMINFO = ^QCMINFO;
  TQCMInfo = QCMINFO;

  // For the  SFVM_GETBUTTONINFO message
const
  {$EXTERNALSYM TBIF_APPEND}
  TBIF_APPEND    = 0;
  {$EXTERNALSYM TBIF_PREPEND}
  TBIF_PREPEND   = 1;
  {$EXTERNALSYM TBIF_REPLACE}
  TBIF_REPLACE   = 2;
  {$EXTERNALSYM TBIF_DEFAULT}
  TBIF_DEFAULT   = 0;
  {$EXTERNALSYM TBIF_INTERNETBAR}
  TBIF_INTERNETBAR = $00010000;
  {$EXTERNALSYM TBIF_STANDARDTOOLBAR}
  TBIF_STANDARDTOOLBAR = $00020000;
  {$EXTERNALSYM TBIF_NOTOOLBAR}
  TBIF_NOTOOLBAR  = $00030000;

type
  TBINFO = record
    cbuttons : UINT;       // out
    uFlags : UINT;         // out (one of TBIF_ flags)
  end;
  TTBInfo = TBINFO;
  PTBInfo = ^TBINFO;

//------------------------------------------------------------------------------
// Catagory Interfaces
//------------------------------------------------------------------------------

const
  {$EXTERNALSYM SID_ICategorizer}
  SID_ICategorizer = '{A3B14589-9174-49A8-89A3-06A1AE2B9BA7}';

  {$EXTERNALSYM SID_ICategoryProvider}
  SID_ICategoryProvider = '{9AF64809-5864-4C26-A720-C1F78C086EE3}';

  {$EXTERNALSYM IID_ICategorizer}
  IID_ICategorizer: TGUID = (
    D1:$A3B14589; D2:$9174; D3:$49A8; D4:($89,$A3,$06,$A1,$AE,$2B,$9B,$A7));
  {$EXTERNALSYM IID_ICategoryProvider}
  IID_ICategoryProvider: TGUID = (
    D1:$9AF64809; D2:$5864; D3:$4C26; D4:($A7,$20,$C1,$F7,$8C,$08,$6E,$E3));
  {$EXTERNALSYM CLSID_DefCategoryProvider}

  CLSID_DefCategoryProvider: TGUID = (
    D1:$B2F2E083; D2:$84FE; D3:$4a7e; D4:($80,$C3,$4B,$50,$D1,$0D,$64,$6E));
  {$EXTERNALSYM CLSID_AlphabeticalCategorizer}
  CLSID_AlphabeticalCategorizer: TGUID = (
    D1:$3C2654C6; D2:$7372; D3:$4F6B; D4:($B3,$10,$55,$D6,$12,$8F,$49,$D2));
  {$EXTERNALSYM CLSID_DriveSizeCategorizer}
  CLSID_DriveSizeCategorizer: TGUID = (
    D1:$94357B53; D2:$CA29; D3:$4B78; D4:($83,$AE,$E8,$FE,$74,$09,$13,$4F));
  {$EXTERNALSYM CLSID_DriveTypeCategorizer}
  CLSID_DriveTypeCategorizer: TGUID = (
    D1:$B0A8F3CF; D2:$4333; D3:$4BAB; D4:($88,$73,$1C,$CB,$1C,$AD,$A4,$8B));
  {$EXTERNALSYM CLSID_FreeSpaceCategorizer}
  CLSID_FreeSpaceCategorizer: TGUID = (
    D1:$B5607793; D2:$24AC; D3:$44C7; D4:($82,$E2,$83,$17,$26,$AA,$6C,$B7));
  {$EXTERNALSYM CLSID_SizeCategorizer}
  CLSID_SizeCategorizer: TGUID = (
    D1:$55D7B852; D2:$F6D1; D3:$42F2; D4:($AA,$75,$87,$28,$A1,$B2,$D2,$64));
  {$EXTERNALSYM CLSID_TimeCategorizer}
  CLSID_TimeCategorizer: TGUID = (
    D1:$3BB4118F; D2:$DDFD; D3:$4D30; D4:($A3,$48,$9F,$B5,$D6,$BF,$1A,$FE));
  {$EXTERNALSYM CLSID_MergedCategorizer}
  CLSID_MergedCategorizer: TGUID = (
    D1:$8E827C11; D2:$33E7; D3:$4BC1; D4:($B2,$42,$8C,$D9,$A1,$C2,$B3,$04));

  {$EXTERNALSYM CATINFO_NORMAL}
  CATINFO_NORMAL = $0000;
  {$EXTERNALSYM CATINFO_COLLAPSED}
  CATINFO_COLLAPSED = $00000001;
  {$EXTERNALSYM CATINFO_HIDDEN}
  CATINFO_HIDDEN = $00000002;

  {$EXTERNALSYM CATSORT_DEFAULT}
  CATSORT_DEFAULT = $00000000;
  {$EXTERNALSYM CATSORT_NAME}
  CATSORT_NAME = $00000001;

const
  // These are the possibilites for the TSHColumnInfo structure
  {$EXTERNALSYM FMTID_Storage}
  FMTID_Storage: TGUID = (
    D1:$B725F130; D2:$47EF; D3:$101A; D4:($A5,$F1,$02,$60,$8C,$9E,$EB,$AC));
  {$EXTERNALSYM PID_STG_NAME}
  PID_STG_NAME = 10;        // The object's display name VT_BSTR
  {$EXTERNALSYM PID_STG_STORAGETYPE}
  PID_STG_STORAGETYPE = 4;  // The object's type VT_BSTR
  {$EXTERNALSYM PID_STG_SIZE}
  PID_STG_SIZE = 12;        // The object's size VT_BSTR
  {$EXTERNALSYM PID_STG_WRITETIME}
  PID_STG_WRITETIME = 14;   // The object's modified attribute VT_BSTR
  {$EXTERNALSYM PID_STG_ATTRIBUTES}
  PID_STG_ATTRIBUTES = 13;  // The object's attributes VT_BSTR

  {$EXTERNALSYM FMTID_ShellDetails}
  FMTID_ShellDetails: TGUID = (
    D1:$28636AA6; D2:$953D; D3:$11D2; D4:($B5,$D6,$00,$C0,$4F,$D9,$18,$D0));
  {$EXTERNALSYM PID_FINDDATA}
  PID_FINDDATA = 0;         // A WIN32_FIND_DATAW structure. VT_ARRAY | VT_UI1
  {$EXTERNALSYM PID_NETRESOURCE}
  PID_NETRESOURCE = 1;      // A NETRESOURCE structure. VT_ARRAY | VT_UI1
  {$EXTERNALSYM PID_DESCRIPTIONID}
  PID_DESCRIPTIONID = 2;    // A SHDESCRIPTIONID structure. VT_ARRAY | VT_UI1

  {$EXTERNALSYM FMTID_Displaced}
  FMTID_Displaced: TGUID = (
    D1:$9B174B33; D2:$40FF; D3:$11D2; D4:($A2,$7E,$00,$C0,$4F,$C3,$08,$71));
  {$EXTERNALSYM PID_DISPLACED_FROM}
  PID_DISPLACED_FROM = 2;
  {$EXTERNALSYM PID_DISPLACED_DATE}
  PID_DISPLACED_DATE = 3;

  {$EXTERNALSYM FMTID_Misc}
  FMTID_Misc: TGUID = (
    D1:$9B174B35; D2:$40FF; D3:$11D2; D4:($A2,$7E,$00,$C0,$4F,$C3,$08,$71));
  {$EXTERNALSYM PID_MISC_STATUS}
  PID_MISC_STATUS = 2;           // The synchronization status.
  {$EXTERNALSYM PID_MISC_ACCESSCOUNT}
  PID_MISC_ACCESSCOUNT = 3;      // Not used.
  {$EXTERNALSYM PID_MISC_OWNER}
  PID_MISC_OWNER = 4;            // Ownership of the file (for the NTFS file system).
         
  {$EXTERNALSYM FMTID_Volume}
  FMTID_Volume: TGUID = (
    D1:$49691C90; D2:$7E17; D3:$101A; D4:($A9,$1C,$08,$00,$2B,$2E,$CD,$A9));
  {$EXTERNALSYM PID_VOLUME_FREE}
  PID_VOLUME_FREE = 2;           // The amount of free space.

  {$EXTERNALSYM FMTID_Query}
  FMTID_Query: TGUID = (
    D1:$49691C90; D2:$7E17; D3:$101A; D4:($A9,$1C,$08,$00,$2B,$2E,$CD,$A9));
  {$EXTERNALSYM PID_QUERY_RANK}
  PID_QUERY_RANK = 2;            // The rank of the file.

  {$EXTERNALSYM FMTID_SummaryInformation}
  FMTID_SummaryInformation: TGUID = (
    D1:$F29F85E0; D2:$4FF9; D3:$1068; D4:($AB,$91,$08,$00,$2B,$27,$B3,$D9));
  {$EXTERNALSYM PIDSI_TITLE}
  PIDSI_TITLE         = 2;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_SUBJECT}
  PIDSI_SUBJECT       = 3;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_AUTHOR}
  PIDSI_AUTHOR        = 4;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_KEYWORDS}
  PIDSI_KEYWORDS      = 5;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_COMMENTS}
  PIDSI_COMMENTS      = 6;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_TEMPLATE}
  PIDSI_TEMPLATE      = 7;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_LASTAUTHOR}
  PIDSI_LASTAUTHOR    = 8;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_REVNUMBER}
  PIDSI_REVNUMBER     = 9;   // VT_LPSTR
  {$EXTERNALSYM PIDSI_EDITTIME}
  PIDSI_EDITTIME      = 10;  // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_LASTPRINTED}
  PIDSI_LASTPRINTED   = 11;  // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_CREATE_DTM}
  PIDSI_CREATE_DTM    = 12;  // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_LASTSAVE_DTM}
  PIDSI_LASTSAVE_DTM  = 13;  // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_PAGECOUNT}
  PIDSI_PAGECOUNT     = 14;  // VT_I4
  {$EXTERNALSYM PIDSI_WORDCOUNT}
  PIDSI_WORDCOUNT     = 15;  // VT_I4
  {$EXTERNALSYM PIDSI_CHARCOUNT}
  PIDSI_CHARCOUNT     = 16;  // VT_I4
  {$EXTERNALSYM PIDSI_THUMBNAIL}
  PIDSI_THUMBNAIL     = 17;  // VT_CF
  {$EXTERNALSYM PIDSI_APPNAME}
  PIDSI_APPNAME       = 18;  // VT_LPSTR
  {$EXTERNALSYM PIDSI_DOC_SECURITY}
  PIDSI_DOC_SECURITY  = 19;  // VT_I4

type
  tagTCategoryInfo = record
    CategoryInfo: DWORD;  // a CATINFO_XXXX constant
    wscName: packed array[0..259] of WideChar;
  end;

  ICategorizer = interface(IUnknown)
  [SID_ICategorizer]
    function GetDescription(pszDesc: LPWSTR; cch: UINT): HRESULT; stdcall;
    function GetCategory(cidl: UINT; apidl: PItemIDList; rgCategoryIds: TWordArray): HRESULT; stdcall;
    function GetCategoryInfo(dwCategoryId: DWORD; var pci: tagTCategoryInfo): HRESULT; stdcall;
    function CompareCategory(csfFlags: DWORD; dwCategoryId1, dwCategoryId2: DWORD): HRESULT; stdcall;
  end;

  ICategoryProvider = interface(IUnknown)
  [SID_ICategoryProvider]
    function CanCategorizeOnSCID(var pscid: TSHColumnID): HRESULT; stdcall;
    function GetDefaultCategory(const pguid: TGUID; var pscid: TSHColumnID): HRESULT; stdcall;
    function GetCategoryForSCID(var pscid: TSHColumnID; var pguid: TGUID): HRESULT; stdcall;
    function EnumCategories(out penum: IEnumGUID): HRESULT; stdcall;
    function GetCategoryName(const pguid: TGUID; pszName: PWideChar; cch: UINT): HRESULT; stdcall;
    function CreateCategory(const pguid: TGUID; const riid: TGUID; out ppv: ICategorizer): HRESULT; stdcall;
  end;

const
  {$EXTERNALSYM FFFP_EXACTMATCH}
  FFFP_EXACTMATCH = 0;
  {$EXTERNALSYM FFFP_NEARESTPARENTMATCH}
  FFFP_NEARESTPARENTMATCH = 1;

  // KNOWN_FOLDER_FLAG
  {$EXTERNALSYM KF_FLAG_CREATE}
  KF_FLAG_CREATE = $00008000;
  {$EXTERNALSYM KF_FLAG_DONT_VERIFY}
  KF_FLAG_DONT_VERIFY = $00004000;
  {$EXTERNALSYM KF_FLAG_DONT_UNEXPAND}
  KF_FLAG_DONT_UNEXPAND = $00002000;
  {$EXTERNALSYM KF_FLAG_NO_ALIAS}
  KF_FLAG_NO_ALIAS = $00001000;
  {$EXTERNALSYM KF_FLAG_INIT}
  KF_FLAG_INIT = $00000800;
  {$EXTERNALSYM KF_FLAG_DEFAULT_PATH}
  KF_FLAG_DEFAULT_PATH = $00000400;
  {$EXTERNALSYM KF_FLAG_NOT_PARENT_RELATIVE}
  KF_FLAG_NOT_PARENT_RELATIVE = $00000200;
  {$EXTERNALSYM KF_FLAG_SIMPLE_IDLIST}
  KF_FLAG_SIMPLE_IDLIST = $00000100;
  {$EXTERNALSYM KF_FLAG_ALIAS_ONLY}
  KF_FLAG_ALIAS_ONLY = $80000000;

  {$EXTERNALSYM KF_REDIRECT_USER_EXCLUSIVE}
  KF_REDIRECT_USER_EXCLUSIVE = $00000001;
  {$EXTERNALSYM KF_REDIRECT_COPY_SOURCE_DACL}
  KF_REDIRECT_COPY_SOURCE_DACL = $00000002;
  {$EXTERNALSYM KF_REDIRECT_OWNER_USER}
  KF_REDIRECT_OWNER_USER = $00000004;
  {$EXTERNALSYM KF_REDIRECT_SET_OWNER_EXPLICIT}
  KF_REDIRECT_SET_OWNER_EXPLICIT = $00000008;
  {$EXTERNALSYM KF_REDIRECT_CHECK_ONLY}
  KF_REDIRECT_CHECK_ONLY = $00000010;
  {$EXTERNALSYM KF_REDIRECT_WITH_UI}
  KF_REDIRECT_WITH_UI = $00000020;
  {$EXTERNALSYM KF_REDIRECT_UNPIN}
  KF_REDIRECT_UNPIN = $00000040;
  {$EXTERNALSYM KF_REDIRECT_PIN}
  KF_REDIRECT_PIN = $00000080;
  {$EXTERNALSYM KF_REDIRECT_COPY_CONTENTS}
  KF_REDIRECT_COPY_CONTENTS = $00000200;
  {$EXTERNALSYM KF_REDIRECT_DEL_SOURCE_CONTENTS}
  KF_REDIRECT_DEL_SOURCE_CONTENTS = $00000400;
  {$EXTERNALSYM KF_REDIRECT_EXCLUDE_ALL_KNOWN_SUBFOLDERS}
  KF_REDIRECT_EXCLUDE_ALL_KNOWN_SUBFOLDERS = $00000800;

  // KF_CATEGORY
  {$EXTERNALSYM KF_CATEGORY_VIRTUAL}
  KF_CATEGORY_VIRTUAL = 1;
  {$EXTERNALSYM KF_CATEGORY_FIXED}
  KF_CATEGORY_FIXED = 2;
  {$EXTERNALSYM KF_CATEGORY_COMMON}
  KF_CATEGORY_COMMON = 3;
  {$EXTERNALSYM KF_CATEGORY_PERUSER}
  KF_CATEGORY_PERUSER = 4;

  // KF_DEFINITION_FLAGS
  {$EXTERNALSYM KFDF_LOCAL_REDIRECT_ONLY}
  KFDF_LOCAL_REDIRECT_ONLY = $00000002;
  {$EXTERNALSYM KFDF_ROAMABLE}
  KFDF_ROAMABLE = $00000004;
  {$EXTERNALSYM KFDF_PRECREATE}
  KFDF_PRECREATE = $00000008;
  {$EXTERNALSYM KFDF_STREAM}
  KFDF_STREAM = $00000010;
  {$EXTERNALSYM KFDF_PUBLISHEXPANDEDPATH}
  KFDF_PUBLISHEXPANDEDPATH = $00000020;

  //KF_REDIRECTION_CAPABILITIES
  {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_ALLOW_ALL}
  KF_REDIRECTION_CAPABILITIES_ALLOW_ALL = $000000FF;
  {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_REDIRECTABLE}
  KF_REDIRECTION_CAPABILITIES_REDIRECTABLE = $00000001;
  {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_DENY_ALL}
  KF_REDIRECTION_CAPABILITIES_DENY_ALL = $000FFF00;
  {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_DENY_POLICY_REDIRECTED}
  KF_REDIRECTION_CAPABILITIES_DENY_POLICY_REDIRECTED = $00000100;
   {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_DENY_POLICY}
  KF_REDIRECTION_CAPABILITIES_DENY_POLICY = $00000200;
  {$EXTERNALSYM KF_REDIRECTION_CAPABILITIES_DENY_PERMISSIONS}
  KF_REDIRECTION_CAPABILITIES_DENY_PERMISSIONS = $00000400;


  SID_IKnownFolderManager = '{8BE2D872-86AA-4d47-B776-32CCA40C7018}' ;
  IID_IKnowndFolderManager: TGUID = (
    D1:$8BE2D872; D2:$86AA; D3:$4d47; D4:($B7,$76,$32,$CC,$A4,$0C,$70,$18));
  SID_IKnownFolder = '{3AA7AF7E-9B36-420c-A8E3-F77D4674A488}' ;
  IID_IKnowndFolder: TGUID = (
    D1:$3AA7AF7E; D2:$9B36; D3:$420c; D4:($A8,$E3,$F7,$7D,$46,$74,$A4,$88));

type
  TKnownFolderIDs = array[0..0] of TGUID;
  PKnownFolderIDs = ^TKnownFolderIDs;

  {$EXTERNALSYM TKnownFolderDefinition}
  TKnownFolderDefinition = record
    Category: DWORD;      // KF_CATEGORY
    pszName: LPWSTR;
    pszDescription: LPWSTR;
    fidParent: TGUID;
    pszRelativePath: LPWSTR;
    pszParsingName: LPWSTR;
    pszTooltip: LPWSTR;
    pszLocalizedName: LPWSTR;
    pszIcon: LPWSTR;
    pszSecurity: LPWSTR;
    dwAttributes: DWORD;
    kfdFlags: DWORD; // KF_DEFINITION_FLAGS
    ftidType: TGUID;
  end;
  PKnownFolderDefinition = ^TKnownFolderDefinition;
  KNOWNFOLDER_DEFINITION = TKnownFolderDefinition;

  {$EXTERNALSYM IKnownFolder}
  IKnownFolder = interface(IUnknown)
  [SID_IKnownFolder]
    function GetCategory(out pCategory: DWORD): HRESULT; stdcall;
    function GetFolderDefinition(pKFD: PKnownFolderDefinition): HRESULT; stdcall;
    function GetFolderType(out ftid: TGUID): HRESULT; stdcall;
    function GetId(out pkfid: TGUID): HRESULT; stdcall;
    function GetIDList(dwFlags: DWORD; out ppidl: PItemIDList): HRESULT; stdcall;
    function GetPath(dwFlags: DWORD; out ppszPath: LPWSTR): HRESULT; stdcall;
    function GetRedirectionCapabilities(out pCapabilities: DWORD): HRESULT; stdcall;
    function GetShellItem(dwFlags: DWORD; const riid: TGUID; var ppv: IUnknown): HRESULT; stdcall;     // IUnknown is typically an IShellItem or IShellItem2
    function SetPath(dwFlags: DWORD; pszPath: LPCWSTR): HRESULT; stdcall;
  end;

  {$EXTERNALSYM TKnownFolderDefinition}
  IKnownFolderManager = interface(IUnknown)
  [SID_IKnownFolderManager]
    function FindFolderFromIDList(pidl: PItemIDList; out ppkf: IKnownFolder): HRESULT; stdcall;
    function FindFolderFromPath(pszPath: LPCWSTR; mode: DWORD; out ppkf: IKnownFolder): HRESULT; stdcall;
    function FolderIdFromCsidl(nCsidl: Integer; out pfid: TGUID): HRESULT; stdcall;
    function FolderIdToCsidl(const rfid: TGUID; var pnCsidl: Integer): HRESULT; stdcall;
    function GetFolder(const rfid: TGUID; out ppkf: IKnownFolder): HRESULT; stdcall;
    function GetFolderByName(pszCanonicalName: LPCWSTR; out ppkf: IKnownFolder): HRESULT; stdcall;
    function GetFolderIds(out ppKFId: PKnownFolderIDs; var pCount: UINT): HRESULT; stdcall;
    function Redirect(const rfid: TGUID; hwnd: HWND; flags: DWORD; pszTargetPath: LPCWSTR; cFolders: UINT; pExclusion: PKnownFolderIDs; out ppszError: LPWSTR): HRESULT; stdcall;
    function RegisterFolder(const rfid: TGUID; pKFD: PKnownFolderDefinition): HRESULT; stdcall;
    function UnregisterFolder(const rfid: TGUID): HRESULT; stdcall;
  end;

const
  {$EXTERNALSYM IID_IBrowserFrameOptions}
  IID_IBrowserFrameOptions : TGUID = (
    D1:$10DF43C8; D2:$1DBE; D3:$11D3; D4:($8B,$34,$00,$60,$97,$DF,$5B,$D4));
  SID_IBrowserFrameOptions  = '{10DF43C8-1DBE-11D3-8B34-006097DF5BD4}';

  // Frame Options
  BFO_BROWSER_PERSIST_SETTINGS          = $0001;
  BFO_RENAME_FOLDER_OPTIONS_TOINTERNET  = $0002;
  BFO_BOTH_OPTIONS                      = $0004;
  BIF_PREFER_INTERNET_SHORTCUT          = $0008;
  BFO_BROWSE_NO_IN_NEW_PROCESS          = $0010;
  BFO_ENABLE_HYPERLINK_TRACKING          = $0020;
  BFO_USE_IE_OFFLINE_SUPPORT            = $0040;
  BFO_SUBSTITUE_INTERNET_START_PAGE      = $0080;
  BFO_USE_IE_LOGOBANDING                = $0100;
  BFO_ADD_IE_TOCAPTIONBAR                = $0200;
  BFO_USE_DIALUP_REF                    = $0400;
  BFO_USE_IE_TOOLBAR                    = $0800;
  BFO_NO_PARENT_FOLDER_SUPPORT          = $1000;
  BFO_NO_REOPEN_NEXT_RESTART            = $2000;
  BFO_GO_HOME_PAGE                      = $4000;
  BFO_PREFER_IEPROCESS                  = $8000;
  BFO_SHOW_NAVIGATION_CANCELLED          = $10000;
  BFO_QUERY_ALL                          = $FFFFFFFF;

type
  TBrowserFrameOption = (
    bfoBrowserPersistSettings,
    bfoRenameFolderOptionsToInternet,
    bfoBothOptions,
    bfoPreferInternetShortcut,
    bfoBrowseNoInNewProcess,
    bfoEnableHyperlinkTracking,
    bfoUseIEOfflineSupport,
    bfoSubstituteInternetStartPage,
    bfoUseIELogoBanding,
    bfoAddIEToCaptionBar,
    bfoUseDialupRef,
    bfoUseIEToolbar,
    bfoNoParentFolderSupport,
    bfoNoReopenNextRestart,
    bfoGoHomePage,
    bfoPreferIEProcess,
    bfoShowNavigationCancelled
  );
  TBrowserFrameOptions = set of TBrowserFrameOption;

const
  bfoNone = [];
  bfoQueryAll  = [bfoBrowserPersistSettings..bfoShowNavigationCancelled];

type
  IBrowserFrameOptions = interface(IUnknown)
  [SID_IBrowserFrameOptions]
    function GetFrameOptions(dwRequested : DWORD; var pdwResult : DWORD) : HResult; stdcall;
  end;

const
  {$EXTERNALSYM IID_IQueryAssociations}
  IID_IQueryAssociations: TGUID = (
    D1:$C46CA590; D2:$3C3F; D3:$11D2; D4:($BE, $E6, $00, $00, $F8, $05, $CA, $57));
  {$EXTERNALSYM SID_IQueryAssociations}
  SID_IQueryAssociations = '{C46CA590-3C3F-11D2-BEE6-0000F805CA57}';

  {$EXTERNALSYM CLSID_QueryAssociations}
  CLSID_QueryAssociations: TGUID = '{A07034FD-6CAA-4954-AC3f-97A27216F98A}';


  {$EXTERNALSYM ASSOCF_INIT_NOREMAPCLSID}
  ASSOCF_INIT_NOREMAPCLSID    = $00000001;   //  do not remap clsids to progids
  {$EXTERNALSYM ASSOCF_INIT_BYEXENAME}
  ASSOCF_INIT_BYEXENAME       = $00000002;   //  executable is being passed in
  {$EXTERNALSYM ASSOCF_OPEN_BYEXENAME}
  ASSOCF_OPEN_BYEXENAME       = $00000002;   //  executable is being passed in
  {$EXTERNALSYM ASSOCF_INIT_DEFAULTTOSTAR}
  ASSOCF_INIT_DEFAULTTOSTAR   = $00000004;   //  treat "*" as theBaseClass
  {$EXTERNALSYM ASSOCF_INIT_DEFAULTTOFOLDER}
  ASSOCF_INIT_DEFAULTTOFOLDER = $00000008;   //  treat "Folder" as the BaseClass
  {$EXTERNALSYM ASSOCF_NOUSERSETTINGS}
  ASSOCF_NOUSERSETTINGS       = $00000010;   //  dont use HKCU
  {$EXTERNALSYM ASSOCF_NOTRUNCATE}
  ASSOCF_NOTRUNCATE           = $00000020;   //  dont truncate the return string
  {$EXTERNALSYM ASSOCF_VERIFY}
  ASSOCF_VERIFY               = $00000040;   //  verify data is accurate (DISK HITS)
  {$EXTERNALSYM ASSOCF_REMAPRUNDLL}
  ASSOCF_REMAPRUNDLL          = $00000080;   //  actually gets info about rundlls target if applicable
  {$EXTERNALSYM ASSOCF_NOFIXUPS}
  ASSOCF_NOFIXUPS             = $00000100;  //  attempt to fix errors if found
  {$EXTERNALSYM ASSOCF_IGNOREBASECLASS}
  ASSOCF_IGNOREBASECLASS      = $00000200;  //  dont recurse into the baseclass

  {$EXTERNALSYM ASSOCDATA_MSIDESCRIPTOR}
  ASSOCDATA_MSIDESCRIPTOR = 1;               //  Component Descriptor to pass to MSI
  {$EXTERNALSYM ASSOCDATA_NOACTIVATEHANDLER}
  ASSOCDATA_NOACTIVATEHANDLER = 2;           //  restrict attempts to activate window
  {$EXTERNALSYM ASSOCDATA_QUERYCLASSSTORE}
  ASSOCDATA_QUERYCLASSSTORE = 3;            //  should check with the NT Class Store
  {$EXTERNALSYM ASSOCDATA_HASPERUSERASSOC}
  ASSOCDATA_HASPERUSERASSOC = 4;            //  defaults to user specified association
  {$EXTERNALSYM ASSOCDATA_EDITFLAGS}
  ASSOCDATA_EDITFLAGS = 5;                  //  Edit flags.
  {$EXTERNALSYM ASSOCDATA_VALUE}
  ASSOCDATA_VALUE = 6;                      //  use pszExtra as the Value name
  {$EXTERNALSYM ASSOCDATA_MAX}
  ASSOCDATA_MAX = 6;                        //  last item in enum...

  {$EXTERNALSYM ASSOCKEY_SHELLEXECCLASS}
  ASSOCKEY_SHELLEXECCLASS = 1;     //  the key that should be passed to
  {$EXTERNALSYM ASSOCKEY_APP}
  ASSOCKEY_APP = 2;                //  the "Application" key for the
  {$EXTERNALSYM ASSOCKEY_CLASS}
  ASSOCKEY_CLASS = 3;              //  the progid or class key
  {$EXTERNALSYM ASSOCKEY_BASECLASS}
  ASSOCKEY_BASECLASS = 4;          //  the BaseClass key
  {$EXTERNALSYM ASSOCKEY_MAX}
  ASSOCKEY_MAX = 4;                //  last item in enum...


  {$EXTERNALSYM ASSOCSTR_COMMAND}
  ASSOCSTR_COMMAND = 1;              //  shell\verb\command string
  {$EXTERNALSYM ASSOCSTR_EXECUTABLE}
  ASSOCSTR_EXECUTABLE = 2;           //  the executable part of command string
  {$EXTERNALSYM ASSOCSTR_FRIENDLYDOCNAME}
  ASSOCSTR_FRIENDLYDOCNAME = 3;      //  friendly name of the document type
  {$EXTERNALSYM ASSOCSTR_FRIENDLYAPPNAME}
  ASSOCSTR_FRIENDLYAPPNAME = 4;     //  friendly name of executable
  {$EXTERNALSYM ASSOCSTR_NOOPEN}
  ASSOCSTR_NOOPEN = 5;              //  noopen value
  {$EXTERNALSYM ASSOCSTR_SHELLNEWVALUE}
  ASSOCSTR_SHELLNEWVALUE = 6;       //  query values under the shellnew key
  {$EXTERNALSYM ASSOCSTR_DDECOMMAND}
  ASSOCSTR_DDECOMMAND = 7;          //  template for DDE commands
  {$EXTERNALSYM ASSOCSTR_DDEIFEXEC}
  ASSOCSTR_DDEIFEXEC = 8;           //  DDECOMMAND to use if just create a process
  {$EXTERNALSYM ASSOCSTR_DDEAPPLICATION}
  ASSOCSTR_DDEAPPLICATION = 9;     //  Application name in DDE broadcast
  {$EXTERNALSYM ASSOCSTR_DDETOPIC}
  ASSOCSTR_DDETOPIC = 10;          //  Topic Name in DDE broadcast
  {$EXTERNALSYM ASSOCSTR_INFOTIP}
  ASSOCSTR_INFOTIP = 11;           //  info tip for an item, or list of
  {$EXTERNALSYM ASSOCSTR_QUICKTIP}
  ASSOCSTR_QUICKTIP = 13;          //  same as ASSOCSTR_INFOTIP, except, this
  {$EXTERNALSYM ASSOCSTR_TILEINFO}
  ASSOCSTR_TILEINFO = 14;          //  similar to ASSOCSTR_INFOTIP - lists important properties for tileview
  {$EXTERNALSYM ASSOCSTR_CONTENTTYPE}
  ASSOCSTR_CONTENTTYPE = 15;       //  MIME Content type
  {$EXTERNALSYM ASSOCSTR_DEFAULTICON}
  ASSOCSTR_DEFAULTICON = 16;       //  Default icon source
  {$EXTERNALSYM ASSOCSTR_SHELLEXTENSION}
  ASSOCSTR_SHELLEXTENSION = 17;    //  Guid string pointing to the Shellex\Shellextensionhandler Value.
  {$EXTERNALSYM ASSOCSTR_MAX}
  ASSOCSTR_MAX = 17;               //  last item in enum...


  {$EXTERNALSYM ASSOCSTR_DDETOPIC}
  ASSOCENUM_NONE = 1;

type
//  {$EXTERNALSYM IQueryAssociations}
  IQueryAssociations = interface(IUnknown)
  [SID_IQueryAssociations]
    function Init(flags: DWORD; pwszAssoc: LPCWSTR; hkProgid: HKEY; hwnd: HWND): HRESULT; stdcall;
    function GetString(flags: DWORD; str: DWORD; pwszExtra: LPCWSTR; pwszOut: LPWSTR; var pcchOut: DWORD): HRESULT; stdcall;
    function GetKey(flags: DWORD; key: DWORD; pwszExtra: LPCWSTR; var phkeyOut: HKEY): HRESULT; stdcall;
    function GetData(flags: DWORD; data: DWORD; pwszExtra: LPCWSTR; out pvOut: Pointer;
      var pcbOut: DWORD): HRESULT; stdcall;
    function GetEnum(flags: DWORD; assocenum: DWORD; pszExtra: LPCWSTR; const riid: TGUID; out ppvOut: Pointer): HRESULT; stdcall;
  end;


{$IFNDEF COMPILER_10_UP}
type
  {$EXTERNALSYM PStartupInfoW}
  PStartupInfoW = ^TStartupInfoW;
  _STARTUPINFOW = record
    cb: DWORD;
    lpReserved: PWideChar;
    lpDesktop: PWideChar;
    lpTitle: PWideChar;
    dwX: DWORD;
    dwY: DWORD;
    dwXSize: DWORD;
    dwYSize: DWORD;
    dwXCountChars: DWORD;
    dwYCountChars: DWORD;
    dwFillAttribute: DWORD;
    dwFlags: DWORD;
    wShowWindow: Word;
    cbReserved2: Word;
    lpReserved2: PByte;
    hStdInput: THandle;
    hStdOutput: THandle;
    hStdError: THandle;
  end;
  {$EXTERNALSYM TStartupInfoW}
  TStartupInfoW = _STARTUPINFOW;
  {$EXTERNALSYM STARTUPINFOW}
  STARTUPINFOW = _STARTUPINFOW;

{$ENDIF}


//------------------------------------------------------------------------------
// ImageList Helper Interfaces
//------------------------------------------------------------------------------

const
  SID_IImageList = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';
  IID_IImageList: TGUID = SID_IImageList;

type
  IImageList = interface(IUnknown)
  [SID_IImageList]
    function Add(Image, Mask: HBITMAP; var Index: Integer): HRESULT; stdcall;
    function ReplaceIcon(IndexToReplace: Integer; Icon: HICON; var Index: Integer): HRESULT; stdcall;
    function SetOverlayImage(iImage: Integer; iOverlay: Integer): HRESULT; stdcall;
    function Replace(Index: Integer; Image, Mask: HBITMAP): HRESULT; stdcall;
    function AddMasked(Image: HBITMAP; MaskColor: COLORREF; var Index: Integer): HRESULT; stdcall;
    function Draw(var DrawParams: TImageListDrawParams): HRESULT; stdcall;
    function Remove(Index: Integer): HRESULT; stdcall;
    function GetIcon(Index: Integer; Flags: UINT; var Icon: HICON): HRESULT; stdcall;
    function GetImageInfo(Index: Integer; var ImageInfo: TImageInfo): HRESULT; stdcall;
    function Copy(iDest: Integer; SourceList: IUnknown; iSource: Integer; Flags: UINT): HRESULT; stdcall;
    function Merge(i1: Integer; List2: IUnknown; i2, dx, dy: Integer; ID: TGUID; out ppvOut): HRESULT; stdcall;
    function Clone(ID: TGUID; out ppvOut): HRESULT; stdcall;
    function GetImageRect(Index: Integer; var rc: TRect): HRESULT; stdcall;
    function GetIconSize(var cx, cy: Integer): HRESULT; stdcall;
    function SetIconSize(cx, cy: Integer): HRESULT; stdcall;
    function GetImageCount(var Count: Integer): HRESULT; stdcall;
    function SetImageCount(NewCount: UINT): HRESULT; stdcall;
    function SetBkColor(BkColor: COLORREF; var OldColor: COLORREF): HRESULT; stdcall;
    function GetBkColor(var BkColor: COLORREF): HRESULT; stdcall;
    function BeginDrag(iTrack, dxHotSpot, dyHotSpot: Integer): HRESULT; stdcall;
    function EndDrag: HRESULT; stdcall;
    function DragEnter(hWndLock: HWND; x, y: Integer): HRESULT; stdcall;
    function DragLieave(hWndLock: HWND): HRESULT; stdcall;
    function DragMove(x, y: Integer): HRESULT; stdcall;
    function SetDragCursorImage(Image: IUnknown; iDrag, dxHotSpot, dyHotSpot: Integer): HRESULT; stdcall;
    function DragShowNoLock(fShow: BOOL): HRESULT; stdcall;
    function GetDragImage(var CurrentPos, HotSpot: TPoint; ID: TGUID; out ppvOut): HRESULT; stdcall;
    function GetImageFlags(i: Integer; dwFlags: DWORD): HRESULT; stdcall;
    function GetOverlayImage(iOverlay: Integer; var iIndex: Integer): HRESULT; stdcall;
  end;


const
  {$EXTERNALSYM SHIL_LARGE}
  SHIL_LARGE         = 0;   // normally 32x32
  {$EXTERNALSYM SHIL_SMALL}
  SHIL_SMALL         = 1;   // normally 16x16
  {$EXTERNALSYM SHIL_EXTRALARGE}
  SHIL_EXTRALARGE    = 2;   // normall 48x48
  {$EXTERNALSYM SHIL_SYSSMALL}
  SHIL_SYSSMALL      = 3;   // like SHIL_SMALL, but tracks system small icon metric correctly
  {$EXTERNALSYM SHIL_JUMBO}
  SHIL_JUMBO         = 4;   // 256x256, Vista and later only
  {$EXTERNALSYM SHIL_LAST}
  SHIL_LAST          = SHIL_SYSSMALL;

type
  TSHGetImageList = function(iImageList: Integer; const RefID: TGUID; out ImageList): HRESULT; stdcall;


  // Add new CommCtl v6.0 paramters
  {$EXTERNALSYM _IMAGELISTDRAWPARAMS}
  _IMAGELISTDRAWPARAMS = record
    cbSize: DWORD;
    himl: HIMAGELIST;
    i: Integer;
    hdcDst: HDC;
    x: Integer;
    y: Integer;
    cx: Integer;
    cy: Integer;
    xBitmap: Integer;        // x offest from the upperleft of bitmap
    yBitmap: Integer;        // y offset from the upperleft of bitmap
    rgbBk: COLORREF;
    rgbFg: COLORREF;
    fStyle: UINT;
    dwRop: DWORD;
    fState: DWORD;         // CommCtl 6.0 and up
    Frame: DWORD;          // CommCtl 6.0 and up
    crEffect: COLORREF;    // CommCtl 6.0 and up
  end;
  PImageListDrawParams = ^TImageListDrawParams;
  TImageListDrawParams = _IMAGELISTDRAWPARAMS;

const
  {$EXTERNALSYM ILS_NORMAL}
  ILS_NORMAL = $0000;   //The image state is not modified.
  {$EXTERNALSYM ILS_GLOW}
  ILS_GLOW = $0001;     // Not supported.
  {$EXTERNALSYM ILS_SHADOW}
  ILS_SHADOW = $0002;   // Not supported.
  {$EXTERNALSYM ILS_SATURATE}
  ILS_SATURATE = $0002; //Reduces the color saturation of the icon to grayscale. This only affects 32bpp images.
  {$EXTERNALSYM ILS_ALPHA}
  ILS_ALPHA = $0008;    //Alpha blends the icon. Alpha blending controls the transparency level of an icon, according to the value of its alpha channel. The value of the alpha channel is indicated by the frame member in the IMAGELISTDRAWPARAMS method. The alpha channel can be from 0 to 255, with 0 being completely transparent, and 255 being completely opaque.

  {$EXTERNALSYM ILIF_ALPHA}
  ILIF_ALPHA = $0001;   // Set if ImageList is a 32Bit image using GetItemFlags

//------------------------------------------------------------------------------

//------------------------------------------------------------------------------
// Definitions for CDefFolderMenu_Create2
//------------------------------------------------------------------------------
const
  // Extra undocumented callback messages
  // (discovered by Maksym Schipka and Henk Devos)
  {$EXTERNALSYM DFM_MERGECONTEXTMENU}
  DFM_MERGECONTEXTMENU = 1;      // uFlags       LPQCMINFO
  {$EXTERNALSYM DFM_INVOKECOMMAND}
  DFM_INVOKECOMMAND = 2;         // idCmd        pszArgs
  {$EXTERNALSYM DFM_CREATE}
  DFM_CREATE = 3;                // AddRef?
  {$EXTERNALSYM DFM_DESTROY}
  DFM_DESTROY = 4;               // Release
  {$EXTERNALSYM DFM_GETHELPTEXTA}
  DFM_GETHELPTEXTA = 5;          // idCmd,cchMax pszText
  {$EXTERNALSYM DFM_MEASUREITEM}
  DFM_MEASUREITEM =  6;          // same as WM_MEASUREITEM
  {$EXTERNALSYM DFM_DRAWITEM}
  DFM_DRAWITEM = 7;              // same as WM_DRAWITEM
  {$EXTERNALSYM DFM_INITMENUPOPUP}
  DFM_INITMENUPOPUP = 8;         // same as WM_INITMENUPOPUP
  {$EXTERNALSYM DFM_VALIDATECMD}
  DFM_VALIDATECMD = 9;           // idCmd        0
  {$EXTERNALSYM DFM_MERGECONTEXTMENU_TOP}
  DFM_MERGECONTEXTMENU_TOP = 10; // uFlags       LPQCMINFO
  {$EXTERNALSYM DFM_GETHELPTEXTW}
  DFM_GETHELPTEXTW = 11;         // idCmd,cchMax pszText -Unicode
  {$EXTERNALSYM DFM_INVOKECOMMANDEX}
  DFM_INVOKECOMMANDEX = 12;      // idCmd        PDFMICS
  {$EXTERNALSYM DFM_MAPCOMMANDNAME}
  DFM_MAPCOMMANDNAME = 13;       // idCmd *      pszCommandName
  {$EXTERNALSYM DFM_GETDEFSTATICID}
  DFM_GETDEFSTATICID = 14;       // idCmd *      0
  {$EXTERNALSYM DFM_GETVERBW}
  DFM_GETVERBW = 15;             // idCmd,cchMax pszText -Unicode
  {$EXTERNALSYM DFM_GETVERBA}
  DFM_GETVERBA = 16;             // idCmd,cchMax pszText -Ansi
  {$EXTERNALSYM DFM_MERGECONTEXTMENU_BOTTOM}
  DFM_MERGECONTEXTMENU_BOTTOM = 17; // uFlags       LPQCMINFO

  // Extra command IDs
  // (from Axel Sommerfeldt and Henk Devos)
  {$EXTERNALSYM DFM_CMD_DELETE}
  DFM_CMD_DELETE = cardinal(-1);
  {$EXTERNALSYM DFM_CMD_CUT}
  DFM_CMD_CUT = cardinal(-2);
  {$EXTERNALSYM DFM_CMD_COPY}
  DFM_CMD_COPY = cardinal(-3);
  {$EXTERNALSYM DFM_CMD_CREATESHORTCUT}
  DFM_CMD_CREATESHORTCUT  = cardinal(-4);
  {$EXTERNALSYM DFM_CMD_PROPERTIES}
  DFM_CMD_PROPERTIES  = UINT(-5);
  {$EXTERNALSYM DFM_CMD_NEWFOLDER}
  DFM_CMD_NEWFOLDER	= cardinal(-6);
  {$EXTERNALSYM DFM_CMD_PASTE}
  DFM_CMD_PASTE = cardinal(-7);
  {$EXTERNALSYM DFM_CMD_VIEWLIST}
  DFM_CMD_VIEWLIST	= cardinal(-8);
  {$EXTERNALSYM DFM_CMD_VIEWDETAILS}
  DFM_CMD_VIEWDETAILS	= cardinal(-9);
  {$EXTERNALSYM DFM_CMD_PASTELINK}
  DFM_CMD_PASTELINK = cardinal(-10);
  {$EXTERNALSYM DFM_CMD_PASTESPECIAL}
  DFM_CMD_PASTESPECIAL = cardinal(-11);
  {$EXTERNALSYM DFM_CMD_MODALPROP}
  DFM_CMD_MODALPROP = cardinal(-12);

  {$EXTERNALSYM CMIC_MASK_CONTROL_DOWN}
  CMIC_MASK_CONTROL_DOWN = $40000000;
  {$EXTERNALSYM CMIC_MASK_SHIFT_DOWN}
  CMIC_MASK_SHIFT_DOWN = $10000000;


type

  PDFMICS = ^TDFMICS;
  TDFMICS = record
    cbSize:  DWORD;
    fMask:  DWORD;     // CMIC_MASK_ values for the invoke
    lParam:  LPARAM;  // same as lParam of DFM_INFOKECOMMAND
    idCmdFirst:  UINT;
    idDefMax:  UINT;
    pici:  PCMInvokeCommandInfo; // the whole thing so you can re-invoke on a child
    punkSite: IUnknown;         // site pointer for context menu handler Longhorn and up. Check size to make sure this is valid
  end;

  TFNDFMCallback = function (psf : Ishellfolder; wnd : HWND;
                             pdtObj : IDataObject; uMsg : UINT;
                             WParm : WParam; lParm : LParam) : HResult; stdcall;

//------------------------------------------------------------------------------


//------------------------------------------------------------------------------
// Assorted definitions
//------------------------------------------------------------------------------
const
  {$EXTERNALSYM CMF_EXTENDEDVERBS}
  CMF_EXTENDEDVERBS    =  $0100;

  // Correct definition
  {$EXTERNALSYM SHGetFileInfoW}
  function SHGetFileInfoW(pszPath: PWideChar; dwFileAttributes: DWORD;
    var psfi: TSHFileInfoW; cbFileInfo, uFlags: UINT): DWORD; stdcall;
      external shell32 name 'SHGetFileInfoW';

  {$EXTERNALSYM CreateProcessW}
  function CreateProcessW(lpApplicationName: PWideChar; lpCommandLine: PWideChar;
    lpProcessAttributes, lpThreadAttributes: PSecurityAttributes;
    bInheritHandles: BOOL; dwCreationFlags: DWORD; lpEnvironment: Pointer;
    lpCurrentDirectory: PWideChar; const lpStartupInfo: TStartupInfoW;
    var lpProcessInformation: TProcessInformation): BOOL; stdcall;
      external kernel32 name 'CreateProcessW';

const
  SID_IShellFolderViewCB = '{2047E320-F2A9-11CE-AE65-08002B2E1262}';
  IID_IShellFolderViewCB: TGUID = SID_IShellFolderViewCB;

  {$EXTERNALSYM SFVM_MERGEMENU}
  SFVM_MERGEMENU         =   1;    // <not used> : LPQCMINFO
  {$EXTERNALSYM SFVM_INVOKECOMMAND}
  SFVM_INVOKECOMMAND     =   2;    // idCmd : <not used>
  {$EXTERNALSYM SFVM_GETHELPTEXT}
  SFVM_GETHELPTEXT       =   3;    // idCmd,cchMaxc : pszText
  {$EXTERNALSYM SFVM_GETTOOLTIPTEXT}
  SFVM_GETTOOLTIPTEXT    =   4;    // idCmd,cchMax : pszText
  {$EXTERNALSYM SFVM_GETBUTTONINFO}
  SFVM_GETBUTTONINFO     =   5 ;   // <not used> : LPTBINFO
  {$EXTERNALSYM SFVM_GETBUTTONS}
  SFVM_GETBUTTONS        =   6;    // idCmdFirst,cbtnMax : LPTBBUTTON
  {$EXTERNALSYM SFVM_INITMENUPOPUP}
  SFVM_INITMENUPOPUP     =   7;    // idCmdFirst,nIndex : hmenu
  {$EXTERNALSYM SFVM_SELECTIONCHANGED}
  SFVM_SELECTIONCHANGED   = 8;     // idCmdFirst,nItem : TSFVCBSelectInfo struct
  {$EXTERNALSYM SFVM_DRAWMENUITEM}
  SFVM_DRAWMENUITEM       = 9;     // idCmdFirst : pdis
  {$EXTERNALSYM SFVM_MEASUREMENUITEM}
  SFVM_MEASUREMENUITEM    = 10;    // idCmdFist : pmis
  {$EXTERNALSYM SFVM_EXITMENULOOP}
  // called when context menu exits, not main menu
  SFVM_EXITMENULOOP       = 11;    // 0 : 0
  {$EXTERNALSYM SFVM_VIEWRELEASE}
  // indicates that the IShellView object is being released.
  SFVM_VIEWRELEASE        = 12;    // <not used>: lSelChangeInfo
  {$EXTERNALSYM SFVM_GETNAMELENGTH}
  // Sent when beginning label edit.
  SFVM_GETNAMELENGTH      = 13;    // pidlItem : length
  {$EXTERNALSYM SFVM_FSNOTIFY}
  SFVM_FSNOTIFY          =  14;    // LPCITEMIDLIST* : lEvent
  {$EXTERNALSYM SFVM_WINDOWCREATED}
  SFVM_WINDOWCREATED     =  15;    // hwnd : <not used>
  {$EXTERNALSYM SFVM_WINDOWCLOSING}
  SFVM_WINDOWCLOSING      = 16;    // hwnd : PDVSELCHANGEINFO
  {$EXTERNALSYM SFVM_LISTREFRESHED}
  SFVM_LISTREFRESHED      = 17;    // 0 : lSelChangeInfo
  {$EXTERNALSYM SFVM_WINDOWFOCUSED}
  // Sent to inform us that the list view has received the focus.
  SFVM_WINDOWFOCUSED      = 18;    // 0 : 0
  {$EXTERNALSYM SFVM_KILLFOCUS}
  SFVM_KILLFOCUS          = 19;    // 0 : 0
  {$EXTERNALSYM SFVM_REGISTERCOPYHOOK}
  SFVM_REGISTERCOPYHOOK   = 20;    // 0 : 0
  {$EXTERNALSYM SFVM_COPYHOOKCALLBACK}
  SFVM_COPYHOOKCALLBACK   = 21;    // <not used> : LPCOPYHOOKINFO
  {$EXTERNALSYM SFVM_NOTIFY}
  SFVM_NOTIFY               = 22;  //idFrom LPNOTIFY : <not used>
  {$EXTERNALSYM SFVM_GETDETAILSOF}
  SFVM_GETDETAILSOF      =  23;    // iColumn : DETAILSINFO*
  {$EXTERNALSYM SFVM_COLUMNCLICK}
  SFVM_COLUMNCLICK       =  24;    // iColumn: <not used>
  {$EXTERNALSYM SFVM_QUERYFSNOTIFY}
  SFVM_QUERYFSNOTIFY     =  25;    // <not used> : SHChangeNotifyEntry*
  {$EXTERNALSYM SFVM_DEFITEMCOUNT}
  SFVM_DEFITEMCOUNT      =  26;    // <not used> : UINT*
  {$EXTERNALSYM SFVM_DEFVIEWMODE}
  SFVM_DEFVIEWMODE       =  27;    // <not used> : FOLDERVIEWMODE*
  {$EXTERNALSYM SFVM_UNMERGEMENU}
  SFVM_UNMERGEMENU       =  28;    // <not used> : hmenu
  {$EXTERNALSYM SFVM_ADDINGOBJECT}
  SFVM_ADDINGOBJECT       = 29;    // pidl : PDVSELCHANGEINFO
  {$EXTERNALSYM SFVM_REMOVINGOBJECT}
  SFVM_REMOVINGOBJECT     = 30;    // pidl : PDVSELCHANGEINFO
  {$EXTERNALSYM SFVM_UPDATESTATUSBAR}
  SFVM_UPDATESTATUSBAR   =  31;    // fInitialize : <not used>
  {$EXTERNALSYM SFVM_BACKGROUNDENUM}
  SFVM_BACKGROUNDENUM    =  32;    // -
  {$EXTERNALSYM SFVM_GETCOMMANDDIR}
  SFVM_GETCOMMANDDIR      = 33;
  {$EXTERNALSYM SFVM_GETCOLUMNSTREAM}
  // Get an IStream interface
  SFVM_GETCOLUMNSTREAM    = 34;    // READ/WRITE/READWRITE : IStream
  {$EXTERNALSYM SFVM_CANSELECTALL}
  SFVM_CANSELECTALL       = 35;    // <not used> : lSelChangeInfo              -
  {$EXTERNALSYM SFVM_DIDDRAGDROP}
  SFVM_DIDDRAGDROP       =  36;    // dwEffect : IDataObject*
  {$EXTERNALSYM SFVM_SUPPORTSIDENTITY}
  SFVM_SUPPORTSIDENTITY   = 37;    // 0 : 0
  {$EXTERNALSYM SFVM_ISCHILDOBJECT}
  SFVM_ISCHILDOBJECT      = 38;
  {$EXTERNALSYM SFVM_SETISFV}
  SFVM_SETISFV           =  39;    // <not used> : IShellFolderView*
  {$EXTERNALSYM SFVM_THISIDLIST}
  SFVM_THISIDLIST        =  41;    // <not used> : LPITMIDLIST*
  {$EXTERNALSYM SFVM_GETITEM}
  SFVM_GETITEM            = 42;    // iItem : LPITMIDLIST*
  {$EXTERNALSYM SFVM_SETITEM}
  SFVM_SETITEM            = 43;    // iItem : LPITEMIDLIST
  {$EXTERNALSYM SFVM_INDEXOFITEM}
  SFVM_INDEXOFITEM        = 44;    // *iItem : LPITEMIDLIST
  {$EXTERNALSYM SFVM_FINDITEM}
  SFVM_FINDITEM           = 45;    // *iItem : NM_FINDITEM*
  {$EXTERNALSYM SFVM_WNDMAIN}
  SFVM_WNDMAIN            = 46;    // <not used> : hwndMain
  {$EXTERNALSYM SFVM_ADDPROPERTYPAGES}
  SFVM_ADDPROPERTYPAGES  =  47;    // <not used> : SFVM_PROPPAGE_DATA *
  {$EXTERNALSYM SFVM_BACKGROUNDENUMDONE}
  SFVM_BACKGROUNDENUMDONE=  48;    // <not used> : <not used>
  {$EXTERNALSYM SFVM_GETNOTIFY}
  SFVM_GETNOTIFY         =  49;    // LPITEMIDLIST* : LONG*
  {$EXTERNALSYM SFVM_COLUMNCLICK2}
  SFVM_COLUMNCLICK2       = 50;    // nil : column index
  {$EXTERNALSYM SFVM_STANDARDVIEWS}
  SFVM_STANDARDVIEWS      = 51;    // <not used> : BOOL *
  {$EXTERNALSYM SFVM_REUSEEXTVIEW}
  SFVM_REUSEEXTVIEW       = 52;    // <not used> : BOOL *
  {$EXTERNALSYM SFVM_GETSORTDEFAULTS}
  SFVM_GETSORTDEFAULTS   =  53;    // iDirection : iParamSort
  {$EXTERNALSYM SFVM_GETEMPTYTEXT}
  SFVM_GETEMPTYTEXT       = 54;    // cchMax : pszText
  {$EXTERNALSYM SFVM_GETITEMICONINDEX}
  SFVM_GETITEMICONINDEX   = 55;    // iItem : int *piIcon
  {$EXTERNALSYM SFVM_DONTCUSTOMIZE}
  SFVM_DONTCUSTOMIZE      = 56;    // <not used> : BOOL *pbDontCustomize
  {$EXTERNALSYM SFVM_SIZE}
  SFVM_SIZE              =  57;    // <not used> : <not used>
  {$EXTERNALSYM SFVM_GETZONE}
  SFVM_GETZONE           =  58;    // <not used> : DWORD*
  {$EXTERNALSYM SFVM_GETPANE}
  SFVM_GETPANE           =  59;    // Pane ID : DWORD*
  {$EXTERNALSYM SFVM_ISOWNERDATA}
  SFVM_ISOWNERDATA        = 60;    // ISOWNERDATA : BOOL *
  {$EXTERNALSYM SFVM_GETRANGEOBJECT}
  SFVM_GETRANGEOBJECT     = 61;    // iWhich : ILVRange **
  {$EXTERNALSYM SFVM_CACHEHINT}
  SFVM_CACHEHINT          = 62;    // <not used> : NMLVCACHEHINT *
  {$EXTERNALSYM SFVM_GETHELPTOPIC}
  SFVM_GETHELPTOPIC      =  63;    // <not used> : SFVM_HELPTOPIC_DATA *
  {$EXTERNALSYM SFVM_OVERRIDEITEMCOUNT}
  SFVM_OVERRIDEITEMCOUNT  = 64;    // <not used> : UINT*
  {$EXTERNALSYM SFVM_GETHELPTEXTW}
  SFVM_GETHELPTEXTW       = 65;    // idCmd,cchMax : pszText - unicode
  {$EXTERNALSYM SFVM_GETTOOLTIPTEXTW}
  SFVM_GETTOOLTIPTEXTW    = 66;    // idCmd,cchMax : pszText - unicode
  {$EXTERNALSYM SFVM_GETIPERSISTHISTORY}
  SFVM_GETIPERSISTHISTORY = 67;    // <not used> : IPersistHistory **
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_GETANIMATION      =  68;    // HINSTANCE * : WCHAR *
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_GETHELPTEXTA       = 69;    // idCmd,cchMax : pszText - ansi
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_GETTOOLTIPTEXTA    = 70;    // idCmd,cchMax :      pszText - ansi
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_GETICONOVERLAY     = 71;    // iItem : int iOverlayIndex
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_SETICONOVERLAY     = 72;    // iItem : int * piOverlayIndex
  {$EXTERNALSYM SFVM_GETANIMATION}
  SFVM_ALTERDROPEFFECT    = 73;    // DWORD* : IDataObject*

  // XP messages - not all id'ed yet
  {$EXTERNALSYM SFVM_MESSAGE4A}
  SFVM_MESSAGE4A          = 74;
  {$EXTERNALSYM SFVM_MESSAGE4B}
  SFVM_MESSAGE4B          = 75;
  {$EXTERNALSYM SFVM_MESSAGE4C}
  SFVM_MESSAGE4C          = 76;
  {$EXTERNALSYM SFVM_GET_CUSTOMVIEWINFO}
  SFVM_GET_CUSTOMVIEWINFO = 77;
  {$EXTERNALSYM SFVM_MESSAGE4E}
  SFVM_MESSAGE4E          = 78;
  {$EXTERNALSYM SFVM_ENUMERATEDITEMS}
  SFVM_ENUMERATEDITEMS    = 79;
  {$EXTERNALSYM SFVM_GET_VIEW_DATA}
  SFVM_GET_VIEW_DATA      = 80;
  {$EXTERNALSYM SFVM_MESSAGE51}
  SFVM_MESSAGE51          = 81;
  {$EXTERNALSYM SFVM_GET_WEBVIEW_LAYOUT}
  SFVM_GET_WEBVIEW_LAYOUT = 82;
  {$EXTERNALSYM SFVM_GET_WEBVIEW_CONTENT}
  SFVM_GET_WEBVIEW_CONTENT= 83;
  {$EXTERNALSYM SFVM_GET_WEBVIEW_TASKS}
  SFVM_GET_WEBVIEW_TASKS  = 84;
  {$EXTERNALSYM SFVM_MESSAGE55}
  SFVM_MESSAGE55          = 85;
  {$EXTERNALSYM SFVM_GET_WEBVIEW_THEME}
  SFVM_GET_WEBVIEW_THEME  = 86;
  {$EXTERNALSYM SFVM_MESSAGE57}
  SFVM_MESSAGE57          = 87;
  {$EXTERNALSYM SFVM_MESSAGE58}
  SFVM_MESSAGE58          = 88;
  {$EXTERNALSYM SFVM_MESSAGE59}
  SFVM_MESSAGE59          = 89;
  {$EXTERNALSYM SFVM_MESSAGE5A}
  SFVM_MESSAGE5A          = 90;
  {$EXTERNALSYM SFVM_MESSAGE5B}
  SFVM_MESSAGE5B          = 91;
  {$EXTERNALSYM SFVM_GETDEFERREDVIEWSETTINGS}
  SFVM_GETDEFERREDVIEWSETTINGS = 92;

  {$EXTERNALSYM SFVM_REARRANGE}
  SFVM_REARRANGE = $0001;
  {$EXTERNALSYM SFVM_GETARRANGECOLUMN}
  SFVM_GETARRANGECOLUMN = $0002;
  {$EXTERNALSYM SFVM_ADDOBJECT}
  SFVM_ADDOBJECT = $0003;
  {$EXTERNALSYM SFVM_GETITEMCOUNT}
  SFVM_GETITEMCOUNT = $0004;
  {$EXTERNALSYM SFVM_GETITEMPIDL}
  SFVM_GETITEMPIDL = $0005;
  {$EXTERNALSYM SFVM_REMOVEOBJECT}
  SFVM_REMOVEOBJECT = $0006;
  {$EXTERNALSYM SFVM_UPDATEOBJECT}
  SFVM_UPDATEOBJECT = $0007;
  {$EXTERNALSYM SFVM_SETREDRAW}
  SFVM_SETREDRAW = $0008;
  {$EXTERNALSYM SFVM_GETSELECTEDOBJECTS}
  SFVM_GETSELECTEDOBJECTS = $0009;


  {$EXTERNALSYM SIGDN_NORMALDISPLAY}
  SIGDN_NORMALDISPLAY               = $00000000;
  {$EXTERNALSYM SIGDN_PARENTRELATIVEPARSING}
  SIGDN_PARENTRELATIVEPARSING       = $80018001;
  {$EXTERNALSYM SIGDN_PARENTRELATIVEFORADDRESSBAR}
  SIGDN_PARENTRELATIVEFORADDRESSBAR = $8001c001;
  {$EXTERNALSYM SIGDN_DESKTOPABSOLUTEPARSING}
  SIGDN_DESKTOPABSOLUTEPARSING      = $80028000;
  {$EXTERNALSYM SIGDN_PARENTRELATIVEEDITING}
  SIGDN_PARENTRELATIVEEDITING       = $80031001;
  {$EXTERNALSYM SIGDN_DESKTOPABSOLUTEEDITING}
  SIGDN_DESKTOPABSOLUTEEDITING      = $8004c000;
  {$EXTERNALSYM SIGDN_FILESYSPATH}
  SIGDN_FILESYSPATH                 = $80058000;
  {$EXTERNALSYM SIGDN_URL}
  SIGDN_URL                         = $80068000;

  {$EXTERNALSYM SICHINT_DISPLAY}
  SICHINT_DISPLAY   = $00000000; // iOrder based on display in a folder view
  {$EXTERNALSYM SICHINT_ALLFIELDS}
  SICHINT_ALLFIELDS = $80000000; // exact instance compare
  {$EXTERNALSYM SICHINT_CANONICAL}
  SICHINT_CANONICAL = $10000000; // iOrder based on canonical name (better performance)


  {$EXTERNALSYM SID_IUIElement}
  SID_IUIElement = '{EC6FE84F-DC14-4FBB-889F-EA50FE27FE0F}';
  {$EXTERNALSYM SID_IUICommand}
  SID_IUICommand = '{4026DFB9-7691-4142-B71C-DCF08EA4DD9C}';
  {$EXTERNALSYM SID_IEnumUICommand}
  SID_IEnumUICommand = '{869447DA-9F84-4E2A-B92D-00642DC8A911}';


  {$EXTERNALSYM SVGIO_FLAG_VIEWORDER}
  SVGIO_FLAG_VIEWORDER = $80000000;
  {$EXTERNALSYM SVGIO_CHECKED}
  SVGIO_CHECKED = $00000003;
  {$EXTERNALSYM SVGIO_TYPE_MASK}
  SVGIO_TYPE_MASK = $0000000F;


type
  IShellFolderViewCB = interface(IUnknown)
  [SID_IShellFolderViewCB]
    function MessageSFVCB(uMsg: UINT; WParam: WPARAM; LParam: LPARAM): HRESULT; stdcall;
  end;

  TSFVM_WEBVIEW_CONTENT_DATA = packed record
    l1 : integer;
    l2 : integer;
    pUnk : IUnknown; // IUIElement
    pUnk2 : IUnknown; // IUIElement
    pEnum : IEnumIdList;
  end;
  PSFVM_WEBVIEW_CONTENT_DATA = ^TSFVM_WEBVIEW_CONTENT_DATA;

  TSFVM_WEBVIEW_TASKSECTION_DATA = record
    pEnum : IUnknown; // IEnumUICommand
    pEnum2 : IUnknown; // IEnumUICommand
  end;
  PSFVM_WEBVIEW_TASKSECTION_DATA = ^TSFVM_WEBVIEW_TASKSECTION_DATA;

  TSFVM_WEBVIEW_THEME_DATA = record
    pszTheme : PWideChar;
  end;
  PSFVM_WEBVIEW_THEME_DATA = ^TSFVM_WEBVIEW_THEME_DATA;

  TSFVM_WEBVIEW_LAYOUT_DATA = record
    flags : cardinal;
    pUnk : IUnknown; //IPreview3?
  end;
  PSFVM_WEBVIEW_LAYOUT_DATA = ^TSFVM_WEBVIEW_LAYOUT_DATA;

  PSFVCBSelectInfo = ^TSFVCBSelectInfo;
  TSFVCBSelectInfo = record
    uOldState: DWORD;  // 0
    uNewState:  DWORD; // LVIS_SELECTED, LVIS_FOCUSED,...
    pidl : PItemIdList;
  end;

  // Sent by a drop down toolbar button in the view WM_COMMAND message
  PBDDDATA = ^TBDDDATA;
  TBDDDATA = packed record
    hwndFrom: HWND;
    pva: PVariant;
    dwUnused: DWORD
  end;   

  IShellItem = interface
    ['{43826D1E-E718-42EE-BC55-A1E261C37BFE}']
    function BindToHandler(pbc: IBindCtx; rbhid: TGUID; const riid: TIID; out ppvOut: Pointer): HResult; stdcall;
    function GetParent(var ppsi : IShellItem): HResult; stdcall;
    function GetDisplayName(sigdnName: DWORD; var ppszName: POLESTR): HResult; stdcall;
    function GetAttributes(sfgaoMask: DWORD; var psfgaoAttribs: DWORD): HResult; stdcall;
    function Compare(psi: IShellItem; hint: ULONG; var piOrder: integer): HResult; stdcall;
  end;

  IEnumShellItems = interface(IUnknown)
    ['{4670AC35-34A6-4D2B-B7B6-CD665C6189A5}']
    function Next(celt: UINT; out rgelt: IShellItem; out pceltFetched: UINT): HResult; stdcall;
    function Skip(celt: UINT): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppenum: IEnumShellItems): HResult; stdcall;
  end;

  IShellItemArray = interface(IUnknown)
    ['{90CF20DE-73B4-4AA4-BA7A-82FF310AF24A}']
    function BindToHandler(pbc: IBindCtx; const rbhid: TGUID; const riid: TIID; out ppvOut): HResult; stdcall;
    function GetAttrributes(nEnum: Integer; dwRequested: DWORD; out pdwResult: DWORD): HResult; stdcall;
    function GetCount(out pCount: UINT): HResult; stdcall;
    function GetItemAt(nIndex: uint; out ppItem: IShellItem) : HResult; stdcall;
    function EnumItems(out enumShellItems: IEnumShellItems): HResult; stdcall;
  end;

  IUIElement = interface(IUnknown)
    [SID_IUIElement]
    function get_Name(pItemArray: IShellItemArray; var bstrName: TBStr) : HResult; stdcall;
    function get_Icon(pItemArray: IShellItemArray; var bstrName: TBStr) : HResult; stdcall;
    function get_Tooltip(pItemArray: IShellItemArray; var bstrName: TBStr) : HResult; stdcall;
  end;


  PDetailsInfo =^TDetailsInfo;
  TDetailsInfo = record
    pidl: PItemIDList;
    Fmt: Integer;
    cxChar: Integer;
    str: TStrRet;
    iImage: Integer;
  end;


  TShellViewExProc =  function(psvOuter: IShellview; psf: IShellFolder; hwndMain: HWND; uMsg: UINT; wParam: WPARAM; lParam: LPARAM): HRESULT; stdcall;

  TShellViewCreate = record
    dwSize: DWORD;
    pShellFolder: IShellFolder;
    psvOuter: IShellView;
    pfnCallback: IShellFolderViewCB;
  end;

  TSHCreateShellFolderView = function(var psvcbi: TShellViewCreate; out ppv): HRESULT; stdcall;
  // TSHCreateShellFolderViewEx is too buggy to bother with

const
  {$EXTERNALSYM SID_IPersistIDList}
  SID_IPersistIDList = '{1079ACFC-29BD-11D3-8E0D-00C04F6837D5}';
  {$EXTERNALSYM IID_IPersistIDList}
  IID_IPersistIDList: TGUID = SID_IPersistIDList;

type
  {$IFDEF CPPB_6}
    {$EXTERNALSYM IPersistIDList}
  {$ENDIF}
  IPersistIDList = interface(IPersist)
  [SID_IPersistIDList]
    // sets or gets a fully qualifed idlist for an object
    function SetIDList(pidl : PItemIdList) : HResult; stdcall;
    function GetIDList(var pidl : PItemIdList) : HResult; stdcall;
  end;

const
  {$EXTERNALSYM SID_IBindHost}
  SID_IBindHost = '{FC4801A1-2BA9-11CF-A229-00AA003D7352}';
  {$EXTERNALSYM IID_IBindHost}
  IID_IBindHost: TGUID = SID_IBindHost;

  {$EXTERNALSYM SID_IPersistFreeThreadedObject}
  SID_IPersistFreeThreadedObject = '{C7264BF0-EDB6-11D1-8546-006008059368}';
  {$EXTERNALSYM IID_IPersistFreeThreadedObject}
  IID_IPersistFreeThreadedObject: TGUID = SID_IPersistFreeThreadedObject;

  {$EXTERNALSYM IID_IAssociationArray}
  SID_IAssociationArray = '{3B877E3C-67DE-4F9A-B29B-17D0A1521C6A}';
  {$EXTERNALSYM IID_IAssociationArray}
  IID_IAssociationArray: TGUID = SID_IAssociationArray;

    {$EXTERNALSYM SID_IBindProtocol}
  SID_IBindProtocol = '{79EAC9CD-BAF9-11CE-8C82-00AA004BA90B}';
  {$EXTERNALSYM IID_IBindProtocol}
  IID_IBindProtocol: TGUID = SID_IBindProtocol;

  {$EXTERNALSYM SID_IInternetSecurityManager}
  SID_IInternetSecurityManager = '{79EAC9EE-BAF9-11CE-8C82-00AA004BA90B}';
  {$EXTERNALSYM IID_IInternetSecurityManager}
  IID_IInternetSecurityManager: TGUID = SID_IInternetSecurityManager;

//------------------------------------------------------------------------------
const
  {$EXTERNALSYM CCM_FIRST}
  CCM_FIRST = $2000;
  {$EXTERNALSYM CCM_SETWINDOWTHEME}
  CCM_SETWINDOWTHEME = CCM_FIRST + $000B;
  {$EXTERNALSYM TB_SETWINDOWTHEME}
  TB_SETWINDOWTHEME = CCM_SETWINDOWTHEME;
  {$EXTERNALSYM RBBS_USECHEVRON}
  RBBS_USECHEVRON = $00000200;
  {$EXTERNALSYM RBN_CHEVRONPUSHED}
  RBN_CHEVRONPUSHED = RBN_FIRST - 10;

  {$EXTERNALSYM DBID_DELAYINIT}
  DBID_DELAYINIT =      4;
  {$EXTERNALSYM DBID_FINISHINIT}
  DBID_FINISHINIT =     5;
  {$EXTERNALSYM DBID_SETWINDOWTHEME}
  DBID_SETWINDOWTHEME = 6;
  {$EXTERNALSYM DBID_PERMITAUTOHIDE}
  DBID_PERMITAUTOHIDE = 7;

  //
  // SHCreateThread Flags
  //
  {$EXTERNALSYM CTF_INSIST}
  CTF_INSIST = $0001;
  {$EXTERNALSYM CTF_THREAD_REF}
  CTF_THREAD_REF = $0002;
  {$EXTERNALSYM CTF_PROCESS_REF}
  CTF_PROCESS_REF = $0004;
  {$EXTERNALSYM CTF_COINIT}
  CTF_COINIT = $0008;

// *********************************
// BandSite
// *********************************

type
  // BCB needs these defines.
  BANDSITEINFO = record
    dwMask : DWORD;
    dwState : DWORD;
    dwStyle : DWORD;
  end;
  TBandSiteInfo = BANDSITEINFO;
  PBandSiteInfo = ^TBandSiteInfo;

const
  {$EXTERNALSYM BSID_BANDADDED}
  BSID_BANDADDED         = 0;
  {$EXTERNALSYM BSID_BANDREMOVED}
  BSID_BANDREMOVED       = 1;

  {$EXTERNALSYM BSIM_STATE}
  BSIM_STATE             = $00000001;
  {$EXTERNALSYM BSIM_STYLE}
  BSIM_STYLE             = $00000002;

  {$EXTERNALSYM BSSF_VISIBLE}
  BSSF_VISIBLE           = $00000001;
  {$EXTERNALSYM BSSF_NOTITLE}
  BSSF_NOTITLE           = $00000002;
  {$EXTERNALSYM BSSF_UNDELETEABLE}
  BSSF_UNDELETEABLE      = $00001000;

  {$EXTERNALSYM BSIS_AUTOGRIPPER}
  BSIS_AUTOGRIPPER       = $00000000;
  {$EXTERNALSYM BSIS_NOGRIPPER}
  BSIS_NOGRIPPER         = $00000001;
  {$EXTERNALSYM BSIS_ALWAYSGRIPPER}
  BSIS_ALWAYSGRIPPER     = $00000002;
  {$EXTERNALSYM BSIS_LEFTALIGN}
  BSIS_LEFTALIGN         = $00000004;
  {$EXTERNALSYM BSIS_SINGLECLICK}
  BSIS_SINGLECLICK       = $00000008;
  {$EXTERNALSYM BSIS_NOCONTEXTMENU}
  BSIS_NOCONTEXTMENU     = $00000010;
  {$EXTERNALSYM BSIS_NODROPTARGET}
  BSIS_NODROPTARGET      = $00000020;
  {$EXTERNALSYM BSIS_NOCAPTION}
  BSIS_NOCAPTION         = $00000040;
  {$EXTERNALSYM BSIS_PREFERNOLINEBREAK}
  BSIS_PREFERNOLINEBREAK = $00000080;
  {$EXTERNALSYM BSIS_LOCKED}
  BSIS_LOCKED            = $00000100;

type
  IBandSite = interface
    ['{4CF504B0-DE96-11D0-8B3F-00A0C911E8E5}']
    function AddBand(pUnk: IUnknown): HResult; stdcall;
    function EnumBands(Band: UINT; var pdwBandID: DWord): HResult; stdcall;
    function QueryBand(BandID: dword; var ppstb: IDeskBand;
                       var pdwState: DWORD; pszName : LPWSTR;
                       cchName: integer): HResult; stdcall;
    function SetBandState(dwBandID: DWORD; dwMask: DWORD; dwState: DWORD): HResult; stdcall;
    function RemoveBand(dwBandID: DWORD) : HResult; stdcall;
    function GetBandObject(dwBandID: DWORD; const riid: TIID; out ppv: Pointer): HResult; stdcall;
    function SetBandSiteInfo(const pbsinfo: BANDSITEINFO): HResult; stdcall;
    function GetBandSiteInfo(var pbsInfo: BANDSITEINFO): HResult; stdcall;
  end;

//------------------------------------------------------------------------------
// Assorted definitions
//------------------------------------------------------------------------------
const
  {$EXTERNALSYM SHCIDS_COLUMNMASK}
  SHCIDS_COLUMNMASK = $0000FFFF;

  {$EXTERNALSYM SID_SShellBrowser}
  SID_SShellBrowser = '{000214E2-0000-0000-C000-000000000046}';
  {$EXTERNALSYM SID_SShellDesktop}
  SID_SShellDesktop = '{00021400-0000-0000-C000-000000000046}';
  {$EXTERNALSYM IID_SShellBrowser}
  IID_SShellBrowser: TGUID = SID_SShellBrowser;
  {$EXTERNALSYM IID_SShellDesktop}
  IID_SShellDesktop: TGUID = SID_SShellDesktop;

  {$EXTERNALSYM CWM_GETISHELLBROWSER}
  CWM_GETISHELLBROWSER = WM_USER + 7;


  {$EXTERNALSYM WM_XBUTTONDOWN}
  WM_XBUTTONDOWN   = $020B;
  {$EXTERNALSYM WM_XBUTTONUP}
  WM_XBUTTONUP     = $020C;
  {$EXTERNALSYM WM_XBUTTONDBLCLK}
  WM_XBUTTONDBLCLK = $020D;

  {$EXTERNALSYM VID_LargeIcons}
  VID_LargeIcons: TGUID = '{0057D0E0-3573-11CF-AE69-08002B2E1262}';
  {$EXTERNALSYM VID_SmallIcons}
  VID_SmallIcons: TGUID = '{0E1FA5E0-3573-11CF-AE69-08002B2E1262}';
  {$EXTERNALSYM VID_List}
  VID_List: TGUID = '{137E7700-3573-11CF-AE69-08002B2E1262}';
  {$EXTERNALSYM VID_Details}
  VID_Details: TGUID = '{5984FFE0-28D4-11CF-AE66-08002B2E1262}';
  {$EXTERNALSYM VID_Tile}
  VID_Tile: TGUID = '{65F125E5-7BE1-4810-BA9D-D271C8432CE3}';

  {$EXTERNALSYM SHCONTF_INIT_ON_FIRST_NEXT}
  SHCONTF_INIT_ON_FIRST_NEXT = $0100;   // allow EnumObject() to return before validating enum")
  {$EXTERNALSYM SHCONTF_NETPRINTERSRCH}
  SHCONTF_NETPRINTERSRCH = $0200;   //The client is looking for printers")
  {$EXTERNALSYM SHCONTF_SHAREABLE}
  SHCONTF_SHAREABLE = $0400;   //The client is looking sharable resources (remote shares))
  {$EXTERNALSYM SHCONTF_STORAGE}
  SHCONTF_STORAGE = $0800;   // Include items with accessible storage and their ancestors, including hidden items.
   {$EXTERNALSYM SHCONTF_NAVIGATION_ENUM}
  SHCONTF_NAVIGATION_ENUM = $1000;   // Windows 7 and later. Child folders should provide a navigation enumeration.
   {$EXTERNALSYM SHCONTF_FASTITEMS}
  SHCONTF_FASTITEMS = $2000;   // Windows Vista and later. The calling application is looking for resources that can be enumerated quickly.
   {$EXTERNALSYM SHCONTF_FLATLIST}
  SHCONTF_FLATLIST = $4000;   // Windows Vista and later. Enumerate items as a simple list even if the folder itself is not structured in that way.
   {$EXTERNALSYM SHCONTF_ENABLE_ASYNC}
  SHCONTF_ENABLE_ASYNC = $8000;   // Windows Vista and later. The calling application is monitoring for change notifications. This means that the enumerator does not have to return all results. Items can be reported through change notifications.
   {$EXTERNALSYM SHCONTF_INCLUDESUPERHIDDEN}
  SHCONTF_INCLUDESUPERHIDDEN = $10000;   //Windows 7 and later. Include hidden system items in the enumeration. This value does not include hidden non-system items. (To include hidden non-system items, use SHCONTF_INCLUDEHIDDEN.)


  {$EXTERNALSYM SFGAO_CAPABILITYMASK}
  SFGAO_CAPABILITYMASK = $00000177;
  {$EXTERNALSYM SFGAO_DISPLAYATTRMASK}
  SFGAO_DISPLAYATTRMASK =  $000FC000;
  {$EXTERNALSYM SFGAO_CONTENTSMASK}
  SFGAO_CONTENTSMASK = $80000000;
  {$EXTERNALSYM SFGAO_STORAGECAPMASK}
  SFGAO_STORAGECAPMASK = $70C50008;

  {$IFNDEF COMPILER_7_UP}
    {$EXTERNALSYM BIF_SHAREABLE}
    BIF_SHAREABLE          = $8000;
    {$EXTERNALSYM BIF_BROWSEINCLUDEURLS}
    BIF_BROWSEINCLUDEURLS  = $0080;
    {$EXTERNALSYM BIF_NEWDIALOGSTYLE}
    BIF_NEWDIALOGSTYLE     = $0040;
    {$EXTERNALSYM BIF_USENEWUI}
    BIF_USENEWUI = BIF_NEWDIALOGSTYLE or BIF_EDITBOX;
  {$ENDIF COMPILER_6_UP}

  // ***********************************************************************
  //
  // For use with SHGetKnownFolderPath (or SHGetKnownFolderPath_MP in MPCommonUtilities)
  // This function allow getting the real paths to special folders.  In Vista the
  // TNamespace.NameForParsing will return a pseudo path for thing like Program Files folder.
  // This function will get the real path using the KF_FLAG_DEFAULT_PATH flag.
  //
  // Don't link with this function unless only running on Vista or above.  Use
  // if Assigned(SHGetKnownFolderPath_MP) then
  //    SHGetKnownFolderPath_MP( )
  // else
  //    ..Get the path some other way
  //
  // ***********************************************************************
type
  {$EXTERNALSYM SHGetKnownFolderPath}
  SHGetKnownFolderPath = function(const rfid: TGUID; dwFlags: DWord; hToken: THandle; out ppszPath: PWideChar): HRESULT;
  {$EXTERNALSYM SHGetKnownFolderIDList}
  SHGetKnownFolderIDList = function(const rfid: TGUID; dwFlags: DWord; hToken: THandle; out ppidl: PItemIDList): HRESULT;

const
  // Known Folder Type GUIDs
  
  {$EXTERNALSYM SID_FOLDERTYPEID_Invalid}
  SID_FOLDERTYPEID_Invalid = '{57807898-8c4f-4462-bb63-71042380b109}';
  {$EXTERNALSYM FOLDERTYPEID_Invalid}
  FOLDERTYPEID_Invalid: TGUID = SID_FOLDERTYPEID_Invalid;
  {$EXTERNALSYM SID_FOLDERTYPEID_Generic}
  SID_FOLDERTYPEID_Generic = '{5c4f28b5-f869-4e84-8e60-f11db97c5cc7}';
  {$EXTERNALSYM FOLDERTYPEID_Generic}
  FOLDERTYPEID_Generic: TGUID = SID_FOLDERTYPEID_Generic;
  {$EXTERNALSYM SID_FOLDERTYPEID_GenericSearchResults}
  SID_FOLDERTYPEID_GenericSearchResults = '{7fde1a1e-8b31-49a5-93b8-6be14cfa4943}';
  {$EXTERNALSYM FOLDERTYPEID_GenericSearchResults}
  FOLDERTYPEID_GenericSearchResults: TGUID = SID_FOLDERTYPEID_GenericSearchResults;
  {$EXTERNALSYM SID_FOLDERTYPEID_GenericLibrary}
  SID_FOLDERTYPEID_GenericLibrary = '{5f4eab9a-6833-4f61-899d-31cf46979d49}';
  {$EXTERNALSYM FOLDERTYPEID_GenericLibrary}
  FOLDERTYPEID_GenericLibrary: TGUID = SID_FOLDERTYPEID_GenericLibrary;
  {$EXTERNALSYM SID_FOLDERTYPEID_Documents}
  SID_FOLDERTYPEID_Documents = '{7d49d726-3c21-4f05-99aa-fdc2c9474656}';
  {$EXTERNALSYM FOLDERTYPEID_Documents}
  FOLDERTYPEID_Documents: TGUID = SID_FOLDERTYPEID_Documents;
  {$EXTERNALSYM SID_FOLDERTYPEID_Pictures}
  SID_FOLDERTYPEID_Pictures = '{b3690e58-e961-423b-b687-386ebfd83239}';
  {$EXTERNALSYM FOLDERTYPEID_Pictures}
  FOLDERTYPEID_Pictures: TGUID = SID_FOLDERTYPEID_Pictures;
  {$EXTERNALSYM SID_FOLDERTYPEID_Music}
  SID_FOLDERTYPEID_Music = '{94d6ddcc-4a68-4175-a374-bd584a510b78}';
  {$EXTERNALSYM FOLDERTYPEID_Music}
  FOLDERTYPEID_Music: TGUID = SID_FOLDERTYPEID_Music;
  {$EXTERNALSYM SID_FOLDERTYPEID_Videos}
  SID_FOLDERTYPEID_Videos = '{5fa96407-7e77-483c-ac93-691d05850de8}';
  {$EXTERNALSYM FOLDERTYPEID_Videos}
  FOLDERTYPEID_Videos: TGUID = SID_FOLDERTYPEID_Videos;
  {$EXTERNALSYM SID_FOLDERTYPEID_UserFiles}
  SID_FOLDERTYPEID_UserFiles = '{CD0FC69B-71E2-46e5-9690-5BCD9F57AAB3}';
  {$EXTERNALSYM FOLDERTYPEID_UserFiles}
  FOLDERTYPEID_UserFiles: TGUID = SID_FOLDERTYPEID_UserFiles;
  {$EXTERNALSYM SID_FOLDERTYPID_UsersLibraries}
  SID_FOLDERTYPID_UsersLibraries = '{C4D98F09-6124-4fe0-9942-826416082DA9}';
   {$EXTERNALSYM FOLDERTYPID_UsersLibraries}
  FOLDERTYPID_UsersLibraries: TGUID = SID_FOLDERTYPID_UsersLibraries;
  {$EXTERNALSYM SID_FOLDERTYPEID_OtherUsers}
  SID_FOLDERTYPEID_OtherUsers = '{B337FD00-9DD5-4635-A6D4-DA33FD102B7A}';
  {$EXTERNALSYM FOLDERTYPEID_OtherUsers}
  FOLDERTYPEID_OtherUsers: TGUID = SID_FOLDERTYPEID_OtherUsers;
  {$EXTERNALSYM SID_FOLDERTYPEID_PublishedItem}
  SID_FOLDERTYPEID_PublishedItem = '{7F2F5B96-FF74-41da-AFD8-1C78A5F3AEA2}';
  {$EXTERNALSYM FOLDERTYPEID_PublishedItem}
  FOLDERTYPEID_PublishedItem: TGUID = SID_FOLDERTYPEID_PublishedItem;
 {$EXTERNALSYM SID_FOLDERTYPEID_Communications}
  SID_FOLDERTYPEID_Communications = '{91475fe5-586b-4eba-8d75-d17434b8cdf6}';
  {$EXTERNALSYM FOLDERTYPEID_Communications}
  FOLDERTYPEID_Communications: TGUID = SID_FOLDERTYPEID_Communications;
  {$EXTERNALSYM SID_FOLDERTYPEID_Contacts}
  SID_FOLDERTYPEID_Contacts = '{de2b70ec-9bf7-4a93-bd3d-243f7881d492}';
  {$EXTERNALSYM FOLDERTYPEID_Contacts}
  FOLDERTYPEID_Contacts: TGUID = SID_FOLDERTYPEID_Contacts;
  {$EXTERNALSYM SID_FOLDERTYPEID_StartMenu}
  SID_FOLDERTYPEID_StartMenu = '{ef87b4cb-f2ce-4785-8658-4ca6c63e38c6}';
  {$EXTERNALSYM FOLDERTYPEID_StartMenu}
  FOLDERTYPEID_StartMenu: TGUID = SID_FOLDERTYPEID_StartMenu;
  {$EXTERNALSYM SID_FOLDERTYPEID_RecordedTV}
  SID_FOLDERTYPEID_RecordedTV = '{5557a28f-5da6-4f83-8809-c2c98a11a6fa}';
  {$EXTERNALSYM FOLDERTYPEID_RecordedTV}
  FOLDERTYPEID_RecordedTV: TGUID = SID_FOLDERTYPEID_RecordedTV;
  {$EXTERNALSYM SID_FOLDERTYPEID_SavedGames}
  SID_FOLDERTYPEID_SavedGames = '{d0363307-28cb-4106-9f23-2956e3e5e0e7}';
  {$EXTERNALSYM FOLDERTYPEID_SavedGames}
  FOLDERTYPEID_SavedGames: TGUID = SID_FOLDERTYPEID_SavedGames;
  {$EXTERNALSYM SID_FOLDERTYPEID_OpenSearch}
  SID_FOLDERTYPEID_OpenSearch = '{8faf9629-1980-46ff-8023-9dceab9c3ee3}';
  {$EXTERNALSYM FOLDERTYPEID_OpenSearch}
  FOLDERTYPEID_OpenSearch: TGUID = SID_FOLDERTYPEID_OpenSearch;
  {$EXTERNALSYM SID_FOLDERTYPEID_SearchConnector}
  SID_FOLDERTYPEID_SearchConnector = '{982725ee-6f47-479e-b447-812bfa7d2e8f}';
  {$EXTERNALSYM FOLDERTYPEID_SearchConnector}
  FOLDERTYPEID_SearchConnector: TGUID = SID_FOLDERTYPEID_SearchConnector;
  {$EXTERNALSYM SID_FOLDERTYPEID_GamesFolder}
  SID_FOLDERTYPEID_GamesFolder = '{b689b0d0-76d3-4cbb-87f7-585d0e0ce070}';
  {$EXTERNALSYM FOLDERTYPEID_GamesFolder}
  FOLDERTYPEID_GamesFolder: TGUID = SID_FOLDERTYPEID_GamesFolder;
  {$EXTERNALSYM SID_FOLDERTYPEID_ControlPanelCategory}
  SID_FOLDERTYPEID_ControlPanelCategory = '{de4f0660-fa10-4b8f-a494-068b20b22307}';
  {$EXTERNALSYM FOLDERTYPEID_ControlPanelCategory}
  FOLDERTYPEID_ControlPanelCategory: TGUID = SID_FOLDERTYPEID_ControlPanelCategory;
   {$EXTERNALSYM SID_FOLDERTYPEID_ControlPanelClassic}
  SID_FOLDERTYPEID_ControlPanelClassic = '{0c3794f3-b545-43aa-a329-c37430c58d2a}';
  {$EXTERNALSYM FOLDERTYPEID_ControlPanelClassic}
  FOLDERTYPEID_ControlPanelClassic: TGUID = SID_FOLDERTYPEID_ControlPanelClassic;
  {$EXTERNALSYM SID_FOLDERTYPEID_Printers}
  SID_FOLDERTYPEID_Printers = '{2c7bbec6-c844-4a0a-91fa-cef6f59cfda1}';
  {$EXTERNALSYM FOLDERTYPEID_Printers}
  FOLDERTYPEID_Printers: TGUID = SID_FOLDERTYPEID_Printers;
  {$EXTERNALSYM SID_FOLDERTYPEID_RecycleBin}
  SID_FOLDERTYPEID_RecycleBin = '{d6d9e004-cd87-442b-9d57-5e0aeb4f6f72}';
  {$EXTERNALSYM FOLDERTYPEID_RecycleBin}
  FOLDERTYPEID_RecycleBin: TGUID = SID_FOLDERTYPEID_RecycleBin;
   {$EXTERNALSYM SID_FOLDERTYPEID_SoftwareExplorer}
  SID_FOLDERTYPEID_SoftwareExplorer = '{d674391b-52d9-4e07-834e-67c98610f39d}';
  {$EXTERNALSYM FOLDERTYPEID_SoftwareExplorer}
  FOLDERTYPEID_SoftwareExplorer: TGUID = SID_FOLDERTYPEID_SoftwareExplorer;
  {$EXTERNALSYM SID_FOLDERTYPEID_CompressedFolder}
  SID_FOLDERTYPEID_CompressedFolder = '{80213e82-bcfd-4c4f-8817-bb27601267a9}';
  {$EXTERNALSYM FOLDERTYPEID_CompressedFolder}
  FOLDERTYPEID_CompressedFolder: TGUID = SID_FOLDERTYPEID_CompressedFolder;
  {$EXTERNALSYM SID_FOLDERTYPEID_NetworkExplorer}
  SID_FOLDERTYPEID_NetworkExplorer = '{25CC242B-9A7C-4f51-80E0-7A2928FEBE42}';
  {$EXTERNALSYM FOLDERTYPEID_NetworkExplorer}
  FOLDERTYPEID_NetworkExplorer: TGUID = SID_FOLDERTYPEID_NetworkExplorer;
  {$EXTERNALSYM SID_FOLDERTYPEID_Searches}
  SID_FOLDERTYPEID_Searches = '{0b0ba2e3-405f-415e-a6ee-cad625207853}';
  {$EXTERNALSYM FOLDERTYPEID_Searches}
  FOLDERTYPEID_Searches: TGUID = SID_FOLDERTYPEID_Searches;
  {$EXTERNALSYM SID_FOLDERTYPEID_SearchHome}
  SID_FOLDERTYPEID_SearchHome = '{834d8a44-0974-4ed6-866e-f203d80b3810}';
  {$EXTERNALSYM FOLDERTYPEID_SearchHome}
  FOLDERTYPEID_SearchHome: TGUID = SID_FOLDERTYPEID_SearchHome; 


  // legacy CSIDL value: CSIDL_NETWORK
  {$EXTERNALSYM FOLDERID_NetworkFolder}
  FOLDERID_NetworkFolder: TGUID = (
    D1:$D20BEEC4; D2:$5CA8; D3:$4905; D4:($AE,$3B,$BF,$25,$1E,$A0,$9B,$53));
  {$EXTERNALSYM FOLDERID_ComputerFolder}
  FOLDERID_ComputerFolder: TGUID = (
    D1:$0AC0837C; D2:$BBF8; D3:$452A; D4:($85,$0D,$79,$D0,$8E,$66,$7C,$A7));
  {$EXTERNALSYM FOLDERID_InternetFolder}
  FOLDERID_InternetFolder: TGUID = (
    D1:$4D9F7874; D2:$4E0C; D3:$4904; D4:($96,$7B,$40,$B0,$D2,$0C,$3E,$4B));
  {$EXTERNALSYM FOLDERID_ControlPanelFolder}
  FOLDERID_ControlPanelFolder: TGUID = (
    D1:$82A74AEB; D2:$AEB4; D3:$465C; D4:($A0,$14,$D0,$97,$EE,$34,$6D,$63));
  {$EXTERNALSYM FOLDERID_PrintersFolder}
  FOLDERID_PrintersFolder: TGUID = (
    D1:$76FC4E2D; D2:$D6AD; D3:$4519; D4:($A6,$63,$37,$BD,$56,$06,$81,$85));
  {$EXTERNALSYM FOLDERID_SyncManagerFolder}
  FOLDERID_SyncManagerFolder: TGUID = (
    D1:$43668BF8; D2:$C14E; D3:$49B2; D4:($97,$C9,$74,$77,$84,$D7,$84,$B7));
  {$EXTERNALSYM FOLDERID_SyncSetupFolder}
  FOLDERID_SyncSetupFolder: TGUID = (
    D1:$f214138; D2:$b1d3; D3:$4a90; D4:($bb,$a9,$27,$cb,$c0,$c5,$38,$9a));
  {$EXTERNALSYM FOLDERID_ConflictFolder}
  FOLDERID_ConflictFolder: TGUID = (
    D1:$4bfefb45; D2:$347d; D3:$4006; D4:($a5,$be,$ac,$0c,$b0,$56,$71,$92));
  {$EXTERNALSYM FOLDERID_SyncResultsFolder}
  FOLDERID_SyncResultsFolder: TGUID = (
    D1:$289a9a43; D2:$be44; D3:$4057; D4:($a4,$1b,$58,$7a,$76,$d7,$e7,$f9));
  {$EXTERNALSYM FOLDERID_RecycleBinFolder}
  FOLDERID_RecycleBinFolder: TGUID = (
    D1:$B7534046; D2:$3ECB; D3:$4C18; D4:($BE,$4E,$64,$CD,$4C,$B7,$D6,$AC));
  {$EXTERNALSYM FOLDERID_ConnectionsFolder}
  FOLDERID_ConnectionsFolder: TGUID = (
    D1:$6F0CD92B; D2:$2E97; D3:$45D1; D4:($88,$FF,$B0,$D1,$86,$B8,$DE,$DD));
  {$EXTERNALSYM FOLDERID_Fonts}
  FOLDERID_Fonts: TGUID = (
    D1:$FD228CB7; D2:$AE11; D3:$4AE3; D4:($86,$4C,$16,$F3,$91,$0A,$B8,$FE));
  // legacy CSIDL value:  CSIDL_DESKTOP
  {$EXTERNALSYM FOLDERID_Desktop}
  FOLDERID_Desktop: TGUID = (
    D1:$B4BFCC3A; D2:$DB2C; D3:$424C; D4:($B0,$29,$7F,$E9,$9A,$87,$C6,$41));
  {$EXTERNALSYM FOLDERID_Startup}
  FOLDERID_Startup: TGUID = (
    D1:$B97D20BB; D2:$F46A; D3:$4C97; D4:($BA,$10,$5E,$36,$08,$43,$08,$54));
  {$EXTERNALSYM FOLDERID_Programs}
  FOLDERID_Programs: TGUID = (
    D1:$A77F5D77; D2:$2E2B; D3:$44C3; D4:($A6,$A2,$AB,$A6,$01,$05,$4A,$51));
  {$EXTERNALSYM FOLDERID_StartMenu}
  FOLDERID_StartMenu: TGUID = (
    D1:$625B53C3; D2:$AB48; D3:$4EC1; D4:($BA,$1F,$A1,$EF,$41,$46,$FC,$19));
  {$EXTERNALSYM FOLDERID_Recent}
  FOLDERID_Recent: TGUID = (
    D1:$AE50C081; D2:$EBD2; D3:$438A; D4:($86,$55,$8A,$09,$2E,$34,$98,$7A));
  {$EXTERNALSYM FOLDERID_SendTo}
  FOLDERID_SendTo: TGUID = (
    D1:$8983036C; D2:$27C0; D3:$404B; D4:($8F,$08,$10,$2D,$10,$DC,$FD,$74));
  {$EXTERNALSYM FOLDERID_Documents}
  FOLDERID_Documents: TGUID = (
    D1:$FDD39AD0; D2:$238F; D3:$46AF; D4:($AD,$B4,$6C,$85,$48,$03,$69,$C7));
  {$EXTERNALSYM FOLDERID_Favorites}
  FOLDERID_Favorites: TGUID = (
    D1:$1777F761; D2:$68AD; D3:$4D8A; D4:($87,$BD,$30,$B7,$59,$FA,$33,$DD));
  {$EXTERNALSYM FOLDERID_NetHood}
  FOLDERID_NetHood: TGUID = (
    D1:$C5ABBF53; D2:$E17F; D3:$4121; D4:($89,$00,$86,$62,$6F,$C2,$C9,$73));
  {$EXTERNALSYM FOLDERID_PrintHood}
  FOLDERID_PrintHood: TGUID = (
    D1:$9274BD8D; D2:$CFD1; D3:$41C3; D4:($B3,$5E,$B1,$3F,$55,$A7,$58,$F4));
  {$EXTERNALSYM FOLDERID_Templates}
  FOLDERID_Templates: TGUID = (
    D1:$A63293E8; D2:$664E; D3:$48DB; D4:($A0,$79,$DF,$75,$9E,$05,$09,$F7));
  {$EXTERNALSYM FOLDERID_CommonStartup}
  FOLDERID_CommonStartup: TGUID = (
    D1:$82A5EA35; D2:$D9CD; D3:$47C5; D4:($96,$29,$E1,$5D,$2F,$71,$4E,$6E));
  {$EXTERNALSYM FOLDERID_CommonPrograms}
  FOLDERID_CommonPrograms: TGUID = (
    D1:$0139D44E; D2:$6AFE; D3:$49F2; D4:($86,$90,$3D,$AF,$CA,$E6,$FF,$B8));
  {$EXTERNALSYM FOLDERID_CommonStartMenu}
  FOLDERID_CommonStartMenu: TGUID = (
    D1:$A4115719; D2:$D62E; D3:$491D; D4:($AA,$7C,$E7,$4B,$8B,$E3,$B0,$67));
  {$EXTERNALSYM FOLDERID_PublicDesktop}
  FOLDERID_PublicDesktop: TGUID = (
    D1:$C4AA340D; D2:$F20F; D3:$4863; D4:($AF,$EF,$F8,$7E,$F2,$E6,$BA,$25));
  {$EXTERNALSYM FOLDERID_ProgramData}
  FOLDERID_ProgramData: TGUID = (
    D1:$62AB5D82; D2:$FDC1; D3:$4DC3; D4:($A9,$DD,$07,$0D,$1D,$49,$5D,$97));
  {$EXTERNALSYM FOLDERID_CommonTemplates}
  FOLDERID_CommonTemplates: TGUID = (
    D1:$B94237E7; D2:$57AC; D3:$4347; D4:($91,$51,$B0,$8C,$6C,$32,$D1,$F7));
  {$EXTERNALSYM FOLDERID_PublicDocuments}
  FOLDERID_PublicDocuments: TGUID = (
    D1:$ED4824AF; D2:$DCE4; D3:$45A8; D4:($81,$E2,$FC,$79,$65,$08,$36,$34));
  {$EXTERNALSYM FOLDERID_RoamingAppData}
  FOLDERID_RoamingAppData: TGUID = (
    D1:$3EB685DB; D2:$65F9; D3:$4CF6; D4:($A0,$3A,$E3,$EF,$65,$72,$9F,$3D));
  {$EXTERNALSYM FOLDERID_LocalAppData}
  FOLDERID_LocalAppData: TGUID = (
    D1:$F1B32785; D2:$6FBA; D3:$4FCF; D4:($9D,$55,$7B,$8E,$7F,$15,$70,$91));
  {$EXTERNALSYM FOLDERID_LocalAppDataLow}
  FOLDERID_LocalAppDataLow: TGUID = (
    D1:$A520A1A4; D2:$1780; D3:$4FF6; D4:($BD,$18,$16,$73,$43,$C5,$AF,$16));
  {$EXTERNALSYM FOLDERID_InternetCache}
  FOLDERID_InternetCache: TGUID = (
    D1:$352481E8; D2:$33BE; D3:$4251; D4:($BA,$85,$60,$07,$CA,$ED,$CF,$9D));
  {$EXTERNALSYM FOLDERID_Cookies}
  FOLDERID_Cookies: TGUID = (
    D1:$2B0F765D; D2:$C0E9; D3:$4171; D4:($90,$8E,$08,$A6,$11,$B8,$4F,$F6));
  {$EXTERNALSYM FOLDERID_History}
  FOLDERID_History: TGUID = (
    D1:$D9DC8A3B; D2:$B784; D3:$432E; D4:($A7,$81,$5A,$11,$30,$A7,$59,$63));
  {$EXTERNALSYM FOLDERID_System}
  FOLDERID_System: TGUID = (
    D1:$1AC14E77; D2:$02E7; D3:$4E5D; D4:($B7,$44,$2E,$B1,$AE,$51,$98,$B7));
  {$EXTERNALSYM FOLDERID_SystemX86}
  FOLDERID_SystemX86: TGUID = (
    D1:$D65231B0; D2:$B2F1; D3:$4857; D4:($A4,$CE,$A8,$E7,$C6,$EA,$7D,$27));
  {$EXTERNALSYM FOLDERID_Windows}
  FOLDERID_Windows: TGUID = (
    D1:$F38BF404; D2:$1D43; D3:$42F2; D4:($93,$05,$67,$DE,$0B,$28,$FC,$23));
  {$EXTERNALSYM FOLDERID_Profile}
  FOLDERID_Profile: TGUID = (
    D1:$5E6C858F; D2:$0E22; D3:$4760; D4:($9A,$FE,$EA,$33,$17,$B6,$71,$73));
  {$EXTERNALSYM FOLDERID_Pictures}
  FOLDERID_Pictures: TGUID = (
    D1:$33E28130; D2:$4E1E; D3:$4676; D4:($83,$5A,$98,$39,$5C,$3B,$C3,$BB));
  {$EXTERNALSYM FOLDERID_ProgramFilesX86}
  FOLDERID_ProgramFilesX86: TGUID = (
    D1:$7C5A40EF; D2:$A0FB; D3:$4BFC; D4:($87,$4A,$C0,$F2,$E0,$B9,$FA,$8E));
  {$EXTERNALSYM FOLDERID_ProgramFilesCommonX86}
  FOLDERID_ProgramFilesCommonX86: TGUID = (
    D1:$DE974D24; D2:$D9C6; D3:$4D3E; D4:($BF,$91,$F4,$45,$51,$20,$B9,$17));
  {$EXTERNALSYM FOLDERID_ProgramFilesX64}
  FOLDERID_ProgramFilesX64: TGUID = (
    D1:$6d809377; D2:$6af0; D3:$444b; D4:($89,$57,$a3,$77,$3f,$02,$20,$0e));
  {$EXTERNALSYM FOLDERID_ProgramFilesCommonX64}
  FOLDERID_ProgramFilesCommonX64: TGUID = (
    D1:$6365d5a7; D2:$f0d; D3:$45e5; D4:($87,$f6,$d,$a5,$6b,$6a,$4f,$7d));
  {$EXTERNALSYM FOLDERID_ProgramFiles}
  FOLDERID_ProgramFiles: TGUID = (
    D1:$905e63b6; D2:$c1bf; D3:$494e; D4:($b2,$9c,$65,$b7,$32,$d3,$d2,$1a));
  {$EXTERNALSYM FOLDERID_ProgramFilesCommon}
  FOLDERID_ProgramFilesCommon: TGUID = (
    D1:$F7F1ED05; D2:$9F6D; D3:$47A2; D4:($AA,$AE,$29,$D3,$17,$C6,$F0,$66));
  {$EXTERNALSYM FOLDERID_AdminTools}
  FOLDERID_AdminTools: TGUID = (
    D1:$724EF170; D2:$A42D; D3:$4FEF; D4:($9F,$26,$B6,$0E,$84,$6F,$BA,$4F));
  {$EXTERNALSYM FOLDERID_CommonAdminTools}
  FOLDERID_CommonAdminTools: TGUID = (
    D1:$D0384E7D; D2:$BAC3; D3:$4797; D4:($8F,$14,$CB,$A2,$29,$B3,$92,$B5));
  {$EXTERNALSYM FOLDERID_Music}
  FOLDERID_Music: TGUID = (
    D1:$4BD8D571; D2:$6D19; D3:$48D3; D4:($BE,$97,$42,$22,$20,$08,$0E,$43));
  {$EXTERNALSYM FOLDERID_Videos}
  FOLDERID_Videos: TGUID = (
    D1:$18989B1D; D2:$99B5; D3:$455B; D4:($84,$1C,$AB,$7C,$74,$E4,$DD,$FC));
  {$EXTERNALSYM FOLDERID_PublicPictures}
  FOLDERID_PublicPictures: TGUID = (
    D1:$B6EBFB86; D2:$6907; D3:$413C; D4:($9A,$F7,$4F,$C2,$AB,$F0,$7C,$C5));
  {$EXTERNALSYM FOLDERID_PublicMusic}
  FOLDERID_PublicMusic: TGUID = (
    D1:$3214FAB5; D2:$9757; D3:$4298; D4:($BB,$61,$92,$A9,$DE,$AA,$44,$FF));
  {$EXTERNALSYM FOLDERID_PublicVideos}
  FOLDERID_PublicVideos: TGUID = (
    D1:$2400183A; D2:$6185; D3:$49FB; D4:($A2,$D8,$4A,$39,$2A,$60,$2B,$A3));
  {$EXTERNALSYM FOLDERID_ResourceDir}
  FOLDERID_ResourceDir: TGUID = (
    D1:$8AD10C31; D2:$2ADB; D3:$4296; D4:($A8,$F7,$E4,$70,$12,$32,$C9,$72));
  {$EXTERNALSYM FOLDERID_LocalizedResourcesDir}
  FOLDERID_LocalizedResourcesDir: TGUID = (
    D1:$2A00375E; D2:$224C; D3:$49DE; D4:($B8,$D1,$44,$0D,$F7,$EF,$3D,$DC));
  {$EXTERNALSYM FOLDERID_CommonOEMLinks}
  FOLDERID_CommonOEMLinks: TGUID = (
    D1:$C1BAE2D0; D2:$10DF; D3:$4334; D4:($BE,$DD,$7A,$A2,$0B,$22,$7A,$9D));
  {$EXTERNALSYM FOLDERID_CDBurning}
  FOLDERID_CDBurning: TGUID = (
    D1:$9E52AB10; D2:$F80D; D3:$49DF; D4:($AC,$B8,$43,$30,$F5,$68,$78,$55));
  {$EXTERNALSYM FOLDERID_UserProfiles}
  FOLDERID_UserProfiles: TGUID = (
    D1:$0762D272; D2:$C50A; D3:$4BB0; D4:($A3,$82,$69,$7D,$CD,$72,$9B,$80));
  {$EXTERNALSYM FOLDERID_Playlists}
  FOLDERID_Playlists: TGUID = (
    D1:$DE92C1C7; D2:$837F; D3:$4F69; D4:($A3,$BB,$86,$E6,$31,$20,$4A,$23));
  {$EXTERNALSYM FOLDERID_SamplePlaylists}
  FOLDERID_SamplePlaylists: TGUID = (
    D1:$15CA69B3; D2:$30EE; D3:$49C1; D4:($AC,$E1,$6B,$5E,$C3,$72,$AF,$B5));
  {$EXTERNALSYM FOLDERID_SampleMusic}
  FOLDERID_SampleMusic: TGUID = (
    D1:$B250C668; D2:$F57D; D3:$4EE1; D4:($A6,$3C,$29,$0E,$E7,$D1,$AA,$1F));
  {$EXTERNALSYM FOLDERID_SamplePictures}
  FOLDERID_SamplePictures: TGUID = (
    D1:$C4900540; D2:$2379; D3:$4C75; D4:($84,$4B,$64,$E6,$FA,$F8,$71,$6B));
  {$EXTERNALSYM FOLDERID_SampleVideos}
  FOLDERID_SampleVideos: TGUID = (
    D1:$859EAD94; D2:$2E85; D3:$48AD; D4:($A7,$1A,$09,$69,$CB,$56,$A6,$CD));
  {$EXTERNALSYM FOLDERID_PhotoAlbums}
  FOLDERID_PhotoAlbums: TGUID = (
    D1:$69D2CF90; D2:$FC33; D3:$4FB7; D4:($9A,$0C,$EB,$B0,$F0,$FC,$B4,$3C));
  {$EXTERNALSYM FOLDERID_Public}
  FOLDERID_Public: TGUID = (
    D1:$DFDF76A2; D2:$C82A; D3:$4D63; D4:($90,$6A,$56,$44,$AC,$45,$73,$85));
  {$EXTERNALSYM FOLDERID_ChangeRemovePrograms}
  FOLDERID_ChangeRemovePrograms: TGUID = (
    D1:$df7266ac; D2:$9274; D3:$4867; D4:($8d,$55,$3b,$d6,$61,$de,$87,$2d));
  {$EXTERNALSYM FOLDERID_AppUpdates}
  FOLDERID_AppUpdates: TGUID = (
    D1:$a305ce99; D2:$f527; D3:$492b; D4:($8b,$1a,$7e,$76,$fa,$98,$d6,$e4));
  {$EXTERNALSYM FOLDERID_AddNewPrograms}
  FOLDERID_AddNewPrograms: TGUID = (
    D1:$de61d971; D2:$5ebc; D3:$4f02; D4:($a3,$a9,$6c,$82,$89,$5e,$5c,$04));
  {$EXTERNALSYM FOLDERID_Downloads}
  FOLDERID_Downloads: TGUID = (
    D1:$374de290; D2:$123f; D3:$4565; D4:($91,$64,$39,$c4,$92,$5e,$46,$7b));
  {$EXTERNALSYM FOLDERID_PublicDownloads}
  FOLDERID_PublicDownloads: TGUID = (
    D1:$3d644c9b; D2:$1fb8; D3:$4f30; D4:($9b,$45,$f6,$70,$23,$5f,$79,$c0));
  {$EXTERNALSYM FOLDERID_SavedSearches}
  FOLDERID_SavedSearches: TGUID = (
    D1:$7d1d3a04; D2:$debb; D3:$4115; D4:($95,$cf,$2f,$29,$da,$29,$20,$da));
  {$EXTERNALSYM FOLDERID_QuickLaunch}
  FOLDERID_QuickLaunch: TGUID = (
    D1:$52a4f021; D2:$7b75; D3:$7b75; D4:($9f,$6b,$4b,$87,$a2,$10,$bc,$8f));
  {$EXTERNALSYM FOLDERID_Contacts}
  FOLDERID_Contacts: TGUID = (
    D1:$56784854; D2:$c6cb; D3:$462b; D4:($81,$69,$88,$e3,$50,$ac,$b8,$82));
  {$EXTERNALSYM FOLDERID_SidebarParts}
  FOLDERID_SidebarParts: TGUID = (
    D1:$a75d362e; D2:$50fc; D3:$4fb7; D4:($ac,$2c,$a8,$be,$aa,$31,$44,$93));
  {$EXTERNALSYM FOLDERID_SidebarDefaultParts}
  FOLDERID_SidebarDefaultParts: TGUID = (
    D1:$7b396e54; D2:$9ec5; D3:$4300; D4:($be,$a,$24,$82,$eb,$ae,$1a,$26));
  {$EXTERNALSYM FOLDERID_TreeProperties}
  FOLDERID_TreeProperties: TGUID = (
    D1:$5b3749ad; D2:$b49f; D3:$49c1; D4:($83,$eb,$15,$37,$0f,$bd,$48,$82));
  {$EXTERNALSYM FOLDERID_PublicGameTasks}
  FOLDERID_PublicGameTasks: TGUID = (
    D1:$debf2536; D2:$e1a8; D3:$4c59; D4:($b6,$a2,$41,$45,$86,$47,$6a,$ea));
  {$EXTERNALSYM FOLDERID_GameTasks}
  FOLDERID_GameTasks: TGUID = (
    D1:$54fae61; D2:$4dd8; D3:$4787; D4:($80,$b6,$9,$2,$20,$c4,$b7,$0));
  {$EXTERNALSYM FOLDERID_SavedGames}
  FOLDERID_SavedGames: TGUID = (
    D1:$4c5c32ff; D2:$bb9d; D3:$43b0; D4:($b5,$b4,$2d,$72,$e5,$4e,$aa,$a4));
  {$EXTERNALSYM FOLDERID_Games}
  FOLDERID_Games: TGUID = (
    D1:$cac52c1a; D2:$b53d; D3:$4edc; D4:($92,$d7,$6b,$2e,$8a,$c1,$94,$34));
  {$EXTERNALSYM FOLDERID_RecordedTV}
  FOLDERID_RecordedTV: TGUID = (
    D1:$bd85e001; D2:$112e; D3:$431e; D4:($98,$3b,$7b,$15,$ac,$09,$ff,$f1));
  {$EXTERNALSYM FOLDERID_SEARCH_MAPI}
  FOLDERID_SEARCH_MAPI: TGUID = (
    D1:$98ec0e18; D2:$2098; D3:$4d44; D4:($86,$44,$66,$97,$93,$15,$a2,$81));
  {$EXTERNALSYM FOLDERID_SEARCH_CSC}
  FOLDERID_SEARCH_CSC: TGUID = (
    D1:$ee32e446; D2:$31ca; D3:$4aba; D4:($81,$4f,$a5,$eb,$d2,$fd,$6d,$5e));
  {$EXTERNALSYM FOLDERID_Links}
  FOLDERID_Links: TGUID = (
    D1:$bfb9d5e0; D2:$c6a9; D3:$404c; D4:($b2,$b2,$ae,$6d,$b6,$af,$49,$68));
  {$EXTERNALSYM FOLDERID_UsersFiles}
  FOLDERID_UsersFiles: TGUID = (
    D1:$f3ce0f7c; D2:$4901; D3:$4acc; D4:($86,$48,$d5,$d4,$4b,$04,$ef,$8f));
  {$EXTERNALSYM FOLDERID_SearchHome}
  FOLDERID_SearchHome: TGUID = (
    D1:$190337d1; D2:$b8ca; D3:$4121; D4:($a6,$39,$6d,$47,$2d,$16,$97,$2a));
  {$EXTERNALSYM FOLDERID_OriginalImages}
  FOLDERID_OriginalImages: TGUID = (
    D1:$2C36C0AA; D2:$5812; D3:$4b87; D4:($bf,$d0,$4c,$d0,$df,$b1,$9b,$39));
  {$EXTERNALSYM FOLDERID_DeviceMetadataStore}
  FOLDERID_DeviceMetadataStore: TGUID = (
    D1:$5CE4A5E9; D2:$E4EB; D3:$479D; D4:($8,$9f,$13,$0c,$02,$88,$61,$55));
   {$EXTERNALSYM FOLDERID_DocumentsLibrary}
  FOLDERID_DocumentsLibrary: TGUID = (
    D1:$7b0db17d; D2:$9cd2; D3:$4a93; D4:($97,$33,$46,$cc,$89,$02,$2e,$7c));
  {$EXTERNALSYM FOLDERID_HomeGroup}
  FOLDERID_HomeGroup: TGUID = (
    D1:$52528A6B; D2:$B9E3; D3:$4add; D4:($b6,$0d,$58,$8c,$2d,$ba,$84,$2d));
  // Windows 7 Known Folder from here on
  // *************************************************************************
  {$EXTERNALSYM FOLDERID_Libraries}
  FOLDERID_Libraries: TGUID = (
    D1:$1B3EA5DC; D2:$B587; D3:$4786; D4:($b4,$ef,$bd,$1d,$c3,$32,$ae,$ae));
 {$EXTERNALSYM FOLDERID_ImplicitAppShortcuts}
  FOLDERID_ImplicitAppShortcuts: TGUID = (
    D1:$bcb5256f; D2:$79f6; D3:$4cee; D4:($b7,$25,$dc,$34,$e4,$02,$fd,$46));
 {$EXTERNALSYM FOLDERID_MusicLibrary}
  FOLDERID_MusicLibrary: TGUID = (
    D1:$2112AB0A; D2:$C86A; D3:$4ffe; D4:($a3,$68,$0d,$e9,$6e,$47,$01,$2e));
 {$EXTERNALSYM FOLDERID_PicturesLibrary}
  FOLDERID_PicturesLibrary: TGUID = (
    D1:$A990AE9F; D2:$A03B; D3:$4e80; D4:($94,$bc,$99,$12,$d7,$50,$41,$04));
 {$EXTERNALSYM FOLDERID_VideosLibrary}
  FOLDERID_VideosLibrary: TGUID = (
    D1:$491E922F; D2:$5643; D3:$4af4; D4:($a7,$eb,$4e,$7a,$13,$8d,$81,$74));
 {$EXTERNALSYM FOLDERID_RecordedTVLibrary}
  FOLDERID_RecordedTVLibrary: TGUID = (
    D1:$1A6FDBA2; D2:$F42D; D3:$4358; D4:($a7,$98,$b7,$4d,$74,$59,$26,$c5));
 {$EXTERNALSYM FOLDERID_PublicLibraries}
  FOLDERID_PublicLibraries: TGUID = (
    D1:$48daf80b; D2:$e6cf; D3:$4f4e; D4:($b8,$00,$0e,$69,$d8,$4e,$e3,$84));
 {$EXTERNALSYM FOLDERID_UserPinned}
  FOLDERID_UserPinned: TGUID = (
    D1:$9e3995ab; D2:$1f9c; D3:$4f13; D4:($b8,$27,$48,$b2,$4b,$6c,$71,$74));
 {$EXTERNALSYM FOLDERID_Ringtones}
  FOLDERID_Ringtones: TGUID = (
    D1:$C870044B; D2:$F49E; D3:$4126; D4:($a9,$c3,$b5,$2a,$1f,$f4,$11,$e8));
 {$EXTERNALSYM FOLDERID_PublicRingtones}
  FOLDERID_PublicRingtones: TGUID = (
    D1:$E555AB60; D2:$153B; D3:$4D17; D4:($9f,$04,$a5,$fe,$99,$fc,$15,$ec));
 {$EXTERNALSYM FOLDERID_UserProgramFiles}
  FOLDERID_UserProgramFiles: TGUID = (
    D1:$5cd7aee2; D2:$2219; D3:$4a67; D4:($b8,$5d,$6c,$9c,$e1,$56,$60,$cb));
 {$EXTERNALSYM FOLDERID_UserProgramFilesCommon}
  FOLDERID_UserProgramFilesCommon: TGUID = (
    D1:$bcbd3057; D2:$ca5c; D3:$4622; D4:($b4,$2d,$bc,$56,$db,$0a,$e5,$16));
 {$EXTERNALSYM FOLDERID_UsersLibraries}
  FOLDERID_UsersLibraries: TGUID = (
    D1:$A302545D; D2:$DEFF; D3:$464b; D4:($ab,$e8,$61,$c8,$64,$8d,$93,$9b));
 // *************************************************************************


type
  TFindFirstFileExW = function(lpFileName: PWideChar; fInfoLevelId: DWORD; var lpFindFileData: TWIN32FindDataW; fSearchOp: DWORD; lpSearchFilter: Pointer; dwAdditionalFlags: DWORD): THandle; stdcall;
  TFindFirstFileExA = function(lpFileName: PChar; fInfoLevelId: DWORD; var lpFindFileData: TWIN32FindDataW; fSearchOp: DWORD; lpSearchFilter: Pointer; dwAdditionalFlags: DWORD): THandle; stdcall;

const
  // _FINDEX_INFO_LEVELS enumerations
  {$EXTERNALSYM FINDEX_INFO_STANDARD}
  FINDEX_INFO_STANDARD = 0;
  {$EXTERNALSYM FINDEX_INFO_MAX_INFO_LEVEL}
  FINDEX_INFO_MAX_INFO_LEVEL = 1;

  // _FINDEX_SEARCH_OPS enumerations
  {$EXTERNALSYM FINDEX_SEARCH_NAMEMATCH}
  FINDEX_SEARCH_NAMEMATCH = 0;
  {$EXTERNALSYM FINDEX_SEARCH_LIMIT_TO_DIRECTORIES}
  FINDEX_SEARCH_LIMIT_TO_DIRECTORIES = 1;
  {$EXTERNALSYM FINDEX_SEARCH_LIMIT_TO_DEVICES}
  FINDEX_SEARCH_LIMIT_TO_DEVICES = 2;
  {$EXTERNALSYM FINDEX_SEARCH_MAX_SEARCH_OP}
  FINDEX_SEARCH_MAX_SEARCH_OP = 3;

const
  {$EXTERNALSYM SID_IPropertyStore}
  SID_IPropertyStore     = '{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}';
  {$EXTERNALSYM SID_IPropertyDescriptionList}
  SID_IPropertyDescriptionList = '{1f9fc1d0-c39b-4b26-817f-011967d3440e}';

type
//  {$EXTERNALSYM _tagpropertykey}    // BCB does not have this defined and needs it
  _tagpropertykey = record
    fmtid: TGUID;
    pid: DWORD;
  end;
  {$EXTERNALSYM PROPERTYKEY}
  PROPERTYKEY = _tagpropertykey;
  PPropertyKey = ^TPropertyKey;
  TPropertyKey = _tagpropertykey;

//  {$EXTERNALSYM IPropertyStore}  // BCB does not have this defined and needs it
  IPropertyStore = interface(IUnknown)
    [SID_IPropertyStore]
    function GetCount(out cProps: DWORD): HResult; stdcall;
    function GetAt(iProp: DWORD; out pkey: TPropertyKey): HResult; stdcall;
    function GetValue(const key: TPropertyKey; out pv: TPropVariant): HResult; stdcall;
    function SetValue(const key: TPropertyKey; const propvar: TPropVariant): HResult; stdcall;
    function Commit: HResult; stdcall;
  end;

  {$EXTERNALSYM IPropertyDescriptionList}
  IPropertyDescriptionList = interface(IUnknown)
    [SID_IPropertyDescriptionList]
    function GetCount(out pcElem: UINT): HResult; stdcall;
    function GetAt(iElem: UINT; const riid: TIID; out ppv): HResult; stdcall;
  end;

const
  {$EXTERNALSYM FIND_FIRST_EX_CASE_SENSITIVE}
  FIND_FIRST_EX_CASE_SENSITIVE = $0001;

// Header Theme States for Vista
  {$EXTERNALSYM HIS_SORTEDNORMAL}
  HIS_SORTEDNORMAL = 4;
  {$EXTERNALSYM HIS_SORTEDHOT}
	HIS_SORTEDHOT = 5;
  {$EXTERNALSYM HIS_SORTEDPRESSED}
	HIS_SORTEDPRESSED = 6;
  {$EXTERNALSYM HIS_ICONNORMAL}
	HIS_ICONNORMAL = 7;
  {$EXTERNALSYM HIS_ICONHOT}
	HIS_ICONHOT = 8;
  {$EXTERNALSYM HIS_ICONPRESSED}
	HIS_ICONPRESSED = 9;
  {$EXTERNALSYM HIS_ICONSORTEDNORMAL}
	HIS_ICONSORTEDNORMAL = 10;
  {$EXTERNALSYM HIS_ICONSORTEDHOT}
	HIS_ICONSORTEDHOT = 11;
  {$EXTERNALSYM HIS_ICONSORTEDPRESSED}
	HIS_ICONSORTEDPRESSED = 12;

const
    // Library Save Flags
  {$EXTERNALSYM LSF_FAILIFTHERE}
  LSF_FAILIFTHERE        = $00000000;
  {$EXTERNALSYM LSF_OVERRIDEEXISTING}
  LSF_OVERRIDEEXISTING   = $00000001;
  {$EXTERNALSYM LSF_MAKEUNIQUENAME}
  LSF_MAKEUNIQUENAME     = $00000002;

  // Default Save Folder Type
  {$EXTERNALSYM DSFT_DETECT}
  DSFT_DETECT  = $00000001;
   {$EXTERNALSYM DSFT_PRIVATE}
  DSFT_PRIVATE = $00000002;
   {$EXTERNALSYM DSFT_PUBLIC}
  DSFT_PUBLIC  = $00000003;

  // Library Folder Filter
  {$EXTERNALSYM LFF_FORCEFILESYSTEM}
  LFF_FORCEFILESYSTEM  = $00000001;
  {$EXTERNALSYM LFF_STORAGEITEMS}
  LFF_STORAGEITEMS = $00000002;
  {$EXTERNALSYM LFF_ALLITEMS}
  LFF_ALLITEMS  = $00000003;

  // Library Option Flags and Masks
  {$EXTERNALSYM LOF_DEFAULT}
  LOF_DEFAULT           = $00000000;
  {$EXTERNALSYM LOF_PINNEDTONAVPANE}
  LOF_PINNEDTONAVPANE   = $00000001;
  {$EXTERNALSYM LOF_MASK_ALL}
  LOF_MASK_ALL          = $00000001;


  {$IFNDEF COMPILER_14_UP}
type
  {$EXTERNALSYM IShellLibrary}
  IShellLibrary = interface(IUnknown)
  [SID_IShellLibrary]
     function LoadLibraryFromItem(psiLibrary: IShellItem; grfMode: DWORD): HRESULT; stdcall;
     function LoadLibraryFromKonwnFolder(const kfidLibrary: TGUID; grfMode: DWORD): HRESULT; stdcall;
     function AddFolder(psiLocation: IShellItem): HRESULT; stdcall;
     function RemoveFolder(psiLocation: IShellItem): HRESULT; stdcall;
     function GetFolders(lff: DWORD; const riid: TGUID; var psiLocation: IShellItem): HRESULT; stdcall;
     function ResolveFolder(psiFolderToResolve: IShellItem; dwTimeOut: DWORD; const riid: TGUID; var ppv: IUnknown): HRESULT; stdcall;
     function GetDefaultSaveFolder(dsft: DWORD; const riid: TGUID; ppv: IUnknown): HRESULT; stdcall;
     function SetDefaultSaveFolder(dsft: DWORD; psi: IShellItem): HRESULT; stdcall;
     function GetOptions(var plofOptions: DWORD): HRESULT; stdcall;
     function SetOptions(loMask: DWORD; loOptions: DWORD): HRESULT; stdcall;
     function GetFolderType(var pftid: TGUID): HRESULT; stdcall;
     function SetFolderType(const pftid: TGUID): HRESULT; stdcall;
     function GetIcon(out ppszIcon: LPWSTR): HRESULT; stdcall;
     function SetIcon(pszIcon: LPCWSTR): HRESULT; stdcall;
     function Commit: HRESULT; stdcall;
     function Save(psiFolderToSaveIn: IShellItem; pszLibraryName: LPCWSTR; lsf: DWORD; out ppsiSavedTo: IShellItem): HRESULT; stdcall;
     function SaveInKnownFolder(const dfidToSaveIn: TGUID; pszLibraryName: LPCWSTR; lsf: DWORD; out ppsiSavedTo): HRESULT; stdcall;
  end;
  {$ENDIF}

type
  TSHSetThreadRef = function(punk: IUnknown): HRESULT; stdcall;
  TSHGetThreadRef = function(out ppunk: IUnknown): HRESULT; stdcall;
  TSHCreateThreadRef = function(var pcRef: LONGWORD; out ppunk: IUnknown): HRESULT; stdcall;
  TSHReleaseThreadRef = function: HRESULT; stdcall;
  TSHCreateThread = function(pfnThreadProc: TThreadStartRoutine; pData: Pointer;
    dwFlags: DWORD; pfnCallback: TThreadStartRoutine): BOOL; stdcall;
  TSHSetInstanceExplorer = procedure(unk : IUnknown); stdcall;
  TSHGetInstanceExplorer = function(out ppunk: IUnknown): HRESULT; stdcall;


  TWNetGetResourceInformationA = function(lpNetResource: PNetResourceA;
                                      lpBuffer: Pointer; var cbBuffer: DWORD;
                                      var lplpSystem: PAnsiChar): DWORD; stdcall;
  TWNetGetResourceInformationW = function(lpNetResource: PNetResourceW;
                                      lpBuffer: Pointer; var cbBuffer: DWORD;
                                      var lplpSystem: PWideChar): DWORD; stdcall;



const
  shlwapi = 'shlwapi.dll';
  
{$EXTERNALSYM CTL_CODE}
function CTL_CODE(DeviceType, AFunction, Method, Access: Integer): Integer;

implementation

function CTL_CODE(DeviceType, AFunction, Method, Access: Integer): Integer;
begin
   Result:=(DeviceType SHL 16) or (Access SHL 14) or (AFunction SHL 2) or Method;
end;

end.














