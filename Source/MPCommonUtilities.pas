unit MPCommonUtilities;

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

// The following are implemented in Win 95:
//  EnumResourceLanguagesW
//  EnumResourceNamesW
//  EnumResourceTypesW
//  ExtTextOutW
//  FindResourceW
//  FindResourceExW
//  GetCharWidthW
//  GetCommandLineW
//  GetTextExtentPoint32W
//  GetTextExtentPointW
//  lstrlenW
//  MessageBoxExW
//  MessageBoxW
//  MultiByteToWideChar
//  TextOutW
//  WideCharToMultiByte

{$B-}

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  ImgList,
  Math,
  ActiveX,
  ShlObj,
  Variants,
  RTLConsts,
  ShellAPI,
  ComCtrls,
  ComObj,
  Types,
  CommCtrl,
  MPShellTypes,
  MPResources;

const
  WideNull = WideChar(#0);
  WideCR = WideChar(#13);
  WideLF = WideChar(#10);
  WideLineSeparator = WideChar(#2028);
  WideSpace = WideChar(#32);
  WidePeriod = WideChar('.');

  Shlwapi = 'shlwapi.dll';
  Mpr = 'Mpr.dll';


// Global user settable variables
var
  Wow64Enabled: boolean = False;  // Enables a 32 bit app to "see" 64bit system files, it will cause other strange bugs if true


var
  SEasyNSEMsg_Caption: string = 'Shell extension registration';
  SEasyNSEMsg_CannotRegister: string = 'Cannot register shell extension.';
  SEasyNSEMsg_CannotUnRegister: string = 'Cannot unregister shell extension.';
  SEasyNSEMsg_CannotFindRegSvr: string = 'Unable to find RegSvr32.exe executable.';
  SEasyNSEMsg_CannotFindDLL: string = 'Unable to find extension DLL.';

type
  TCommonWideCharArray = array of WideChar;
  TCommonPWideCharArray = array of PWideChar;
  TCommonStringDynArray = array of string;
  TCommonIntegerDynArray = array of Integer;

  TCommonDropEffect = (
    cdeNone,                 // No drop effect (the circle with the slash through it
    cdeCopy,                 // Copy the dropped object
    cdeMove,                 // Move the dropped object
    cdeLink,                 // Make a shortcut to the dropped object
    cdeScroll                // The dragging is in the middle of a scroll
  );
  TCommonDropEffects = set of TCommonDropEffect;

  TCommonOLEDragResult = (
    cdrDrop,                 // The drag resulted in a drop
    cdrCancel,               // The drag resulted in being canceled
    cdrError                 // The drag resulted in an unknown error
  );

  TCommonKeyState = (
    cksControl,              // Control Key is down
    cksLButton,              // Left Mouse is down
    cksMButton,              // Middle Mouse is down
    cksRButton,              // Right Mouse is down
    cksShift,                // Shift Key is down
    cksAlt,                  // Alt Key is down
    cksButton                // One of the mouse buttons are down
  );
  TCommonKeyStates = set of TCommonKeyState;

  TCommonMouseButton = (
    cmbNone,          // No Button
    cmbLeft,          // Left Button
    cmbRight,         // Right Button
    cmbMiddle         // Middle Button
  );

  TCommonVAlignment = (
    cvaTop,                            // The vertical alignment of the text is at the top of the object
    cvaBottom,                         // The vertical alignment of the text is at the bottom of the object
    cvaCenter                          // The vertical alignment of the text is at the center of the object
  );

  // Flags for the DrawTextWEx function
  TCommonDrawTextWFlag = (
    dtSingleLine,       // Put Caption on one line
    dtLeft,             // Aligns Text Left
    dtRight,            // Aligns Text Right
    dtCenter,           // Aligns Text Center
    dtTop,              // Vertical Align Text to Bottom of Rect
    dtBottom,           // Vertical Align Text to Bottom of Rect
                        //   Only valid with: dtSingleLine
    dtVCenter,          // Vertical Align Text to Center of Rect
                        //   Only valid with: dtSingleLine
    dtCalcRect,         // Modifies the Rectangle to the size required for the Text does
                        // not draw the text. By default it does not modify the right
                        // edge of the rectangle, it only changes the height to fit
                        // the text, see dtCalcRectAdjR
    dtCalcRectAdjR,     // Modifies the Rectangles right edge for a best fit of the text
                        // Does not increase the width only shortens it,
                        //   Only valid with: dtCalcRect
    dtCalcRectAlign,    // Modifies the rectangle by aligning it with the original
                        // rectangle based on the dtLeft, dtRight, dtCenter flag.
                        // In other words it ensures that if the text won't fit that
                        // only the end of the text is clipped.  For instance if
                        // the text is horz centered the calculation could clip
                        // both ends of the text.  Just using the dtCalcRectAdjR
                        // flag will only stretch the Right edge and the left will
                        // still be clipped. Using this flag will shift the rect
                        // to the edge of the passed rect so that the beginning
                        // of the text is always shown.
                        //   Only valid with: dtCalcRect and dtCalcRectAdjR
    dtEndEllipsis,      // Adds a "..." to the end of the string if it will not fit in
                        // the passed rectangle
    dtWordBreak,        // Breaks the passed string to best fit in the rectangle
                        // The default Characters to break the line are:
                        //      WideSpace                               ( WideChar(#32) )
                        //      WideCR/WideLF sequence or individually  ( WideChar(#13 or #10) )
                        //      WideLineSeparator                       ( WideChar(#2028) )
    dtUserBreakChars,   // The UserBreakChars parameters should be used for defining
                        // what to use to break the passed string into lines
                        //   Only valid with: dtWordBreak
    dtRTLReading,       // Right to Left reading
    dtNoClip            // Do not clip the text in the rectangle
  );
  TCommonDrawTextWFlags = set of TCommonDrawTextWFlag;

  // Describes the mode how to blend pixels.
  TCommonBlendMode = (
    cbmConstantAlpha,         // apply given constant alpha
    cbmPerPixelAlpha,         // use alpha value of the source pixel
    cbmMasterAlpha,           // use alpha value of source pixel and multiply it with the constant alpha value
    cbmConstantAlphaAndColor  // blend the destination color with the given constant color und the constant alpha value
  );

  TShortenStringEllipsis = (
    sseEnd,            // Ellipsis on end of string
    sseFront,          // Ellipsis on begining of string
    sseMiddle,         // Ellipsis in middle of string
    sseFilePathMiddle  // Ellipsis is in middle of string but tries to show the entire filename
  );

  // RGB (red, green, blue) color given in the range [0..1] per component.
  TCommonRGB = record
    R, G, B: Double;
  end;

  // Hue, luminance, saturation color with all three components in the range [0..1]
  // (so hue's 0..360� is normalized to 0..1).
  TCommonHLS = record
    H, L, S: Double;
  end;

  TDrawWindowButtonType = (
    dwbtMinimize,
    dwbtMaximize,
    dwbtRestore,
    dwbtClose
  );

  TValidateDelimiterExt = (
    vdeColon,     // ":"
    vdeSemiColon, // ";"
    vdeComma,     // ","
    vdePipe       // "|"
  );

  TValidateDelimiterWildCard = (
    vdwcAsterisk,
    vdwcPeriod
  );
  TValidateDelimiterWildCardSet = set of TValidateDelimiterWildCard;

// Enhanced library loading functions that reference counts the loading to make
// sure that libraries are freed when using cool controls in COM applications
function CommonLoadLibrary(LibraryName: string): THandle;
function CommonUnloadLibrary(LibraryName: string): Boolean;
procedure CommonUnloadAllLibraries;

function FlipReverseCopyRect(const Flip, Reverse: Boolean; const Bitmap: TBitmap): TBitmap; overload;
procedure FlipReverseCopyRect(const Flip, Reverse: Boolean; R: TRect; const Canvas: TCanvas); overload;

procedure DrawRadioButton(Canvas: TCanvas; Pos: TPoint; Size: Integer; clBackground, clHotBkGnd,
  clLeftOuter, clRightOuter, clLeftInner, clRightInner: TColor; Checked, Enabled, Hot: Boolean);

procedure DrawCheckBox(Canvas: TCanvas; Pos: TPoint; Size: Integer; clBackground, clHotBkGnd,
  clLeftOuter, clRightOuter, clLeftInner, clRightInner: TColor; Checked, Enabled, Hot: Boolean);

procedure DrawWindowButton(Canvas: TCanvas; Pos: TPoint; Size: Integer; ButtonType: TDrawWindowButtonType);

function CheckBounds(Size: Integer; Character: Char = Char($67)): TRect;

function AbsRect(ARect: TRect): TRect;
function AsyncLeftButtonDown: Boolean;
function AsyncRightButtonDown: Boolean;
function AsyncMiddleButtonDown: Boolean;
function AsyncAltDown: Boolean;
function AsyncControlDown: Boolean;
function AsyncShiftDown: Boolean;
function CalcuateFolderSize(FolderPath: string; Recurse: Boolean): Int64;
function CenterRectHorz(OuterRect, InnerRect: TRect): TRect;
function CenterRectInRect(OuterRect, InnerRect: TRect): TRect;
function CenterRectVert(OuterRect, InnerRect: TRect): TRect;
function CommonSupports(const Instance: IUnknown; const IID: TGUID): Boolean; overload;
function CommonSupports(const Instance: IUnknown; const IID: TGUID; out Intf): Boolean; overload;
function CommonSupports(const Instance: TObject; const IID: TGUID): Boolean; overload;
function CommonSupports(const Instance: TObject; const IID: TGUID; out Intf): Boolean; overload;
procedure CreateProcessMP(ExeFile, Parameters, InitalDir: string);
function DiffRectHorz(Rect1, Rect2: TRect): TRect;
function DiffRectVert(Rect1, Rect2: TRect): TRect;
function DiskInDrive(C: AnsiChar): Boolean;
function DragDetectPlus(Handle: HWND; Pt: TPoint): Boolean;
function DrawTextWEx(DC: HDC; Text: string; var lpRect: TRect; Flags: TCommonDrawTextWFlags; MaxLineCount: Integer): Integer;
function EqualWndMethod(A, B: TWndMethod): Boolean;
function FileIconInit(FullInit: BOOL): BOOL; stdcall;
function FindUniqueMenuID(AMenu: HMenu): Cardinal;
function GetMyDocumentsVirtualFolder: PItemIDList;
function HasMMX: Boolean;
function IsAnyMouseButtonDown: Boolean;
function IsMappedDrivePath(const Path: string): Boolean;
function IsRectNull(ARect: TRect): Boolean;
function IsTextTrueType(Canvas: TCanvas): Boolean; overload;
function IsTextTrueType(DC: HDC): Boolean; overload;
function IsUNCPath(const Path: string): Boolean;
function IsUNCPathSyntax(const Path: string): Boolean;
function IsUnicode: Boolean;    // OS supports Unicode functions (basiclly means IsWinNT or XP)
function IsWin2000: Boolean;
function IsWin95_SR1: Boolean;
function IsWinME: Boolean;
function IsWinNT: Boolean;
function IsWinNT4: Boolean;
function IsWinXP: Boolean;
function IsWinVista: Boolean;
function IsWin7: Boolean;
function ModuleFileName(PathOnly: Boolean = True): string;
function PIDLToPath(PIDL: PItemIDList): string;
function SHGetImageList(iImageList: Integer; const RefID: TGUID; out ppvOut): HRESULT; stdcall;
function ShiftStateToKeys(Keys: TShiftState): LongWord;
function ShiftStateToStr(Keys: TShiftState): string;
function ShortenStringEx(DC: HDC; const S: string; Width: Integer; RTL: Boolean; EllipsisPlacement: TShortenStringEllipsis): string;
function ShortenTextW(DC: hDC; TextToShorten: string; MaxSize: Integer): string;
function ShortFileName(const FileName: string): string;
function ShortPath(const Path: string): string;
function Size(cx, cy: Integer): TSize;
function SplitTextW(DC: hDC; TextToSplit: string; MaxWidth: Integer; var Buffer: TCommonWideCharArray; MaxSplits: Integer): Integer;
function StrRetToStr(StrRet: TStrRet; APIDL: PItemIDList): string;
function SystemDirectory: string;
function SysMenuFont: HFONT;
function SysMenuHeight: Integer;
function TextExtentW(Text: PWideChar; Canvas: TCanvas): TSize; overload;
function TextExtentW(Text: PWideChar; DC: hDC): TSize; overload;
function TextExtentW(Text: string; Canvas: TCanvas): TSize; overload;
function TextExtentW(Text: string; Font: TFont): TSize; overload;
function TextTrueExtentsW(Text: string; DC: HDC): TSize;
function UniqueDirName(const ADirPath: string): string;
function UniqueFileName(const AFilePath: string): string;
function WideExpandEnviromentString(EnviromentString: string): string;
function WideExpandEnviromentStringForUser(EnviromentString: string): string;
function WideExtractFileDrive(Path: string; DriveLetterOnly: Boolean = False): string;
function WideExtractFileExt(Path: string; StripExtPeriod: Boolean = False): string;
function WideExtractFileName(Path: string; BaseNameOnly: Boolean = False): string;
function WideGetTempDir: string;
function WideIncrementalSearch(CompareStr, Mask: string): Integer;
function WideIsDrive(Drive: string): Boolean;
function WideIsFloppy(FileFolder: string): boolean;
function WideIsPathDelimiter(const S: string; Index: Integer): Boolean;
function WideNewFolderName(ParentFolder: string; SuggestedFolderName: string = ''): string;
function WidePathMatchSpec(Path, Mask: string): Boolean;
function WidePathMatchSpecExists: Boolean;
function WideShellExecute(hWnd: HWND; Operation, FileName, Parameters, Directory: string; ShowCmd: Integer = SW_NORMAL): HINST;
function WideStrComp(Str1, Str2: PWideChar): Integer;
function WideStrIComp(Str1, Str2: PWideChar): Integer;
function WideStripExt(AFile: string): string;
function WideStripLeadingBackslash(const S: string): string;
function WideStripRemoteComputer(const UNCPath: string): string;
function WideStripTrailingBackslash(const S: string; Force: Boolean = False): string;
function WideValidateDelimitedExtList(DelimitedText: string; Prefix: TValidateDelimiterWildCardSet; Delimiter: TValidateDelimiterExt): string;
function WindowsDirectory: string;
procedure AlphaBlend(Source, Destination: HDC; R: TRect; Target: TPoint; Mode: TCommonBlendMode; ConstantAlpha, Bias: Integer);
procedure ConvertBitmapEx(Image32: TBitmap; var OutImage: TBitmap; const BackGndColor: TColor);
procedure ShadowBlendBits(Bits: TBitmap; BackGndColor: TColor);
procedure FreeMemAndNil(var P: Pointer);
procedure LoadWideString(S: TStream; var Str: string);
procedure MinMax(var A, B: Integer);
procedure SaveWideString(S: TStream; Str: string);
function UsesAlphaChannel(Image32: TBitmap): Boolean;
procedure WideShowMessage(Window: HWND; ACaption, AMessage: string);

// Rectangle functions
function ProperRect(Rect: TRect): TRect;
function RectHeight(R: TRect): Integer;
function RectToStr(R: TRect): string;
function RectToSquare(R: TRect): TRect;
function RectWidth(R: TRect): Integer;
function ContainsRect(OuterRect, InnerRect: TRect): Boolean;
function IsRectProper(Rect: TRect): Boolean;

// WideString routines (many borrowed from Mike Liscke)
function StrCopyW(Dest, Source: PWideChar): PWideChar;

function KeyDataToShiftState(KeyData: Longint): TShiftState;
function KeyDataAsyncToShiftState: TShiftState;
function KeyToKeyStates(Keys: Word): TCommonKeyStates;
function KeyStatesToMouseButton(Keys: Word): TCommonMouseButton;
function KeyStatesToKey(Keys: TCommonKeyStates): Longword;
function DropEffectToDropEffectStates(Effect: Integer): TCommonDropEffects;
function DropEffectStatesToDropEffect(Effect: TCommonDropEffects): Integer;
function DropEffectToDropEffectState(Effect: Integer): TCommonDropEffect;
function DropEffectStateToDropEffect(Effect: TCommonDropEffect): Integer;
function KeyStateToDropEffect(Keys: TCommonKeyStates): TCommonDropEffect;
function KeyStateToMouseButton(KeyState: TCommonKeyStates): TCommonMouseButton;

// Color Functions
function RGBToHLS(const RGB: TCommonRGB): TCommonHLS;
function HLSToRGB(const HLS: TCommonHLS): TCommonRGB;
function BrightenColor(const RGB: TCommonRGB; Amount: Double): TCommonRGB;
function DarkenColor(const RGB: TCommonRGB; Amount: Double): TCommonRGB;
function MakeTRBG(Color: TColor): TCommonRGB;
function MakeTColor(RGB: TCommonRGB): TColor;
function MakeColorRef(RGB: TCommonRGB; Gamma: Double = 1): COLORREF;
procedure GammaCorrection(var RGB: TCommonRGB; Gamma: Double);
function MakeSafeColor(var RGB: TCommonRGB): Boolean;
function UpsideDownDIB(Bits: TBitmap): Boolean;
procedure FixFormFont(AFont: TFont);

// Window Manipulation
procedure ActivateTopLevelWindow(Child: HWND);
function IsVScrollVisible(Wnd: HWnd): Boolean;
function IsHScrollVisible(Wnd: HWnd): Boolean;
function HScrollbarHeightRuntime(Wnd: HWnd): Integer;
function VScrollbarWidthRuntime(Wnd: HWnd): Integer;
function HScrollbarHeight: Integer;
function VScrollbarWidth: Integer;

// Graphics
procedure FillGradient(X1, Y1, X2, Y2: integer; fStartColor, fStopColor: TColor; StartPoint, EndPoint: integer; fDrawCanvas: TCanvas);
procedure ResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight: Integer);
function ScaleImageList(Source: TCustomImageList; M, D: Integer): TCustomImageList;

// Menu Functions
function AddContextMenuItem(Menu: HMenu; ACaption: string; Index: Integer; MenuID: UINT = $FFFF; hSubMenu: UINT = 0; Enabled: Boolean = True; Checked: Boolean = False; Default: Boolean = False): Integer;
procedure ValidateMenuSeparators(Menu: HMenu);

// Helpers to create a callback function out of a object method
type
  ICallbackStub = interface(IInterface)
    function GetStubPointer: Pointer;
    property StubPointer : Pointer read GetStubPointer;
  end;

  TCallbackStub = class(TInterfacedObject, ICallbackStub)
  private
    fStubPointer : Pointer;
    fCodeSize : integer;
    function GetStubPointer: Pointer;
  public
    constructor Create(Obj : TObject; MethodPtr: Pointer; NumArgs : integer);
    destructor Destroy; override;
  end;

// ****************************************************************************
 // Registration function for Shell and Namespace Extensions
 // Donated by by Alexey Torgashin
 // Note: we assume that paths to both RegSvr32.exe and extension DLL are ANSI (not Unicode).
// ****************************************************************************
type
  TEasyNSERegMessages = set of (enseMsgShowErrors, enseMsgRegSvr);

 //True to register or False to unregister
function RegUnregNSE(const AFileName: string; DoRegister: boolean; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
function RegisterNSE(const AFileName: string; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
function UnregisterNSE(const AFileName: string; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
function ExecShellEx(const Cmd, Params, Dir: string; ShowCmd: integer; DoWait: boolean; WaitForDDE: Boolean = False; WaitForIdleInput: Boolean = False): boolean;

procedure LoadWideFunctions;

//
// Dynamically linked Unicode functions that do not have stubs on Win9x
//
var
  CDefFolderMenu_Create2_MP: function(pidlFolder: PItemIdList; wnd: HWnd; cidl: uint; var apidl: PItemIdList; psf: IShellFolder; lpfn: TFNDFMCallback; nKeys: UINT; ahkeyClsKeys: PHKEY; var ppcm: IContextMenu): HRESULT; stdcall = nil;
  PathMatchSpecW_MP: function(const pszFileParam, pszSpec: PWideChar): Bool; stdcall = nil;
  ExpandEnvironmentStringsForUserW_MP: function(hToken: THandle; lpSrc: PWideChar; lpDst: PWideChar; nSize: DWORD): BOOL; stdcall = nil;
  Wow64RevertWow64FsRedirection_MP: function(OldValue: Pointer): BOOL; stdcall = nil;
  Wow64DisableWow64FsRedirection_MP: function(var OldValue: Pointer): BOOL; stdcall = nil;
  Wow64EnableWow64FsRedirection_MP: function(Wow64FsEnableDirection: BOOLEAN): BOOLEAN; stdcall = nil;
  NetShareEnumW_MP: function(ServerName: PWideChar; Level: DWord; Bufptr: Pointer; Prefmaxlen: DWord; EntriesRead, TotalEntries, resume_handle: LPDWord): DWord; stdcall;


  // Windows 7 Stuff
  SetCurrentProcessExplicitAppUserModelID_MP: function(AppID: PWideChar): HRESULT; stdcall = nil;
  SHGetPropertyStoreForWindow_MP: function(hwnd: HWND; const riid: TGUID; out ppv: IPropertyStore): HRESULT; stdcall;

  NetShareEnum_MP: function(pszServer: PAnsiChar; sLevel: Cardinal; pbBuffer: PAnsiChar; cbBuffer: Cardinal; pcEntriesRead,pcTotalAvail: Pointer): DWord; stdcall;
  NetApiBufferFree_MP: function(Bufptr: Pointer): HRESULT; stdcall = nil;

  // Usage of this flag makes the SumFolder Function not thread safe
  SumFolderAbort: Boolean = False;
  WideFunctionsLoaded: Boolean = False;


implementation

uses
  System.UITypes,
  {$if CompilerVersion >= 21}
  WinCodec,
  {$ifend}
  MPCommonObjects;

type
  PLibRec = ^TLibRec;
  TLibRec = packed record
    LibraryName: string;
    ReferenceCount: Integer;
    Handle: THandle;
  end;

var
  FLibList: TList;
  PIDLMgr: TCommonPIDLManager;
  Shell32Handle,
  Kernel32Handle,
  AdvAPI32Handle,
  UserEnvHandle,
  ShlwapiHandle,
  NetAPI32Handle: THandle;

  UniqueMenuIDSeed: Integer = 1000;

procedure FillSeparatorList(Menu: HMenu; Separators: TList);
var
  i: Integer;
  MenuInfoW: TMenuItemInfoW;
begin
  Separators.Clear;
  for i := 0 to GetMenuItemCount(Menu) - 1 do
  begin
    FillChar(MenuInfoW, SizeOf(MenuInfoW), #0);
    MenuInfoW.cbSize := SizeOf(MenuInfoW);
    MenuInfoW.fMask := MIIM_ID or MIIM_TYPE or MIIM_SUBMENU or MIIM_STATE or MIIM_CHECKMARKS or MIIM_DATA;
    if GetMenuItemInfoW(Menu, i, True, MenuInfoW) then
    begin
      if (MenuInfoW.fType and MFT_SEPARATOR = MFT_SEPARATOR) and (MenuInfoW.hSubMenu = 0) then
        Separators.Add(Pointer(i))
      else
        Separators.Add(Pointer(-1))
    end
  end;
end;

procedure ValidateMenuSeparators(Menu: HMenu);
//
// Removes duplicate separates in a menu
//
var
  i: Integer;
  Separators: TList;
begin
  Separators := TList.Create;
  try
    // Returns a list whose context are $FFFF if the item at index i is a separator
    // else it is the menuItemID
    FillSeparatorList(Menu, Separators);
     // Walk down the list and remove any duplicated separators
    for i := Separators.Count - 1 downto 1 do
    begin
      if (i = GetMenuItemCount(Menu) - 1) then
      begin
        if Integer(Separators[i]) <> -1 then
          DeleteMenu(Menu, i, MF_BYPOSITION)
      end else
      if (Integer(Separators[i]) <> -1) and (Integer(Separators[i-1]) <> -1) then
        DeleteMenu(Menu, i, MF_BYPOSITION)
    end;
    if Separators.Count > 0 then
      begin
        if Integer(Separators[0]) <> -1 then
          DeleteMenu(Menu, 0, MF_BYPOSITION)
      end
  finally
    Separators.Free
  end
end;


function ShiftStateToKeys(Keys: TShiftState): LongWord;
begin
  Result := 0;
  if ssShift in Keys then
    Result := Result or MK_SHIFT;
  if ssCtrl in Keys then
    Result := Result or MK_CONTROL;
  if ssLeft in Keys then
    Result := Result or MK_LBUTTON;
  if ssRight in Keys then
    Result := Result or MK_RBUTTON;
  if ssMiddle in Keys then
    Result := Result or MK_MBUTTON;
  if ssAlt in Keys then
    Result := Result or MK_ALT;
end;

function ShiftStateToStr(Keys: TShiftState): string;
begin
  Result := '[';
  if ssShift in Keys then
    Result := Result + 'ssShift, ';
  if ssCtrl in Keys then
    Result := Result + 'ssCtrl, ';
  if ssLeft in Keys then
    Result := Result + 'ssLeft, ';
  if ssRight in Keys then
    Result := Result + 'ssRight, ';
  if ssMiddle in Keys then
    Result := Result + 'ssMiddle, ';
  if ssAlt in Keys then
    Result := Result + 'ssAlt, ';
  if ssDouble in Keys then
    Result := Result + 'ssDouble, ';
  if Length(Result) > 1 then
    SetLength(Result, Length(Result) - 2);
  Result := Result + ']';
end;

function RegSvrPath: WideString;
const
  ExeName: WideString = 'RegSvr32.exe';
var
  Path: WideString;
begin
  Result:= '';
  //Look in System dir
  Path := SystemDirectory + '\' + ExeName;
  if FileExists(Path) then
    Result := Path
  else begin
    //Look in Windows dir
    Path:= WindowsDirectory + '\' + ExeName;
    if FileExists(Path) then
      Result:= Path;
  end
end;

function ExecShellEx(const Cmd, Params, Dir: string; ShowCmd: integer; DoWait: boolean; WaitForDDE: Boolean = False; WaitForIdleInput: Boolean = False): boolean;
var
  InfoW: TShellExecuteInfoW;
begin
  FillChar(InfoW, SizeOf(InfoW), 0);
  InfoW.cbSize := SizeOf(InfoW);
  InfoW.fMask := SEE_MASK_FLAG_NO_UI or SEE_MASK_NOCLOSEPROCESS;
  if WaitForDDE then
    InfoW.fMask := InfoW.fMask or SEE_MASK_FLAG_DDEWAIT;
  InfoW.lpFile := PWideChar(Cmd);
  InfoW.lpParameters := PWideChar(Params);
  InfoW.lpDirectory := PWideChar(Dir);
  InfoW.nShow := ShowCmd;
  Result := ShellExecuteEx(@InfoW);
  if WaitForIdleInput then
    WaitForInputIdle(InfoW.hProcess, INFINITE);
  if Result and DoWait then
    WaitForSingleObject(InfoW.hProcess, INFINITE);
  if InfoW.hProcess <> 0 then
    CloseHandle(InfoW.hProcess)
end;

function RegUnregNSE(const AFileName: string; DoRegister: boolean; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
var
  Fn, Exe, Msg: WideString;
begin
  Result:= false;

  if not FileExists(AFileName) then
    begin
      if enseMsgShowErrors in AMessages then
      begin
        if DoRegister then
          Msg := SEasyNSEMsg_CannotRegister
        else
          Msg := SEasyNSEMsg_CannotUnRegister;


          MessageBox(Application.Handle,
                   PWideChar( Msg + #13 + SEasyNSEMsg_CannotFindDLL),
                   PWideChar( SEasyNSEMsg_Caption),
                   MB_OK or MB_ICONERROR);
      end;
      Exit
    end;

  Fn := AFileName;
  // If includes spaces then " " it
  if Pos(' ', Fn) > 0 then
    Fn := '"' + Fn + '"';

  Exe := RegSvrPath;

  if not FileExists(Exe) then
  begin
    if enseMsgShowErrors in AMessages then
    begin
      if DoRegister then
        Msg := SEasyNSEMsg_CannotRegister
      else
        Msg := SEasyNSEMsg_CannotUnRegister;

      MessageBoxW(Application.Handle,
                 PWideChar( Msg + #13 + SEasyNSEMsg_CannotFindRegSvr),
                 PWideChar( SEasyNSEMsg_Caption),
                 MB_OK or MB_ICONERROR);
    end;
    Exit
  end;

  if DoRegister then
    Msg := '' else
  Msg := '/U ';
  if not (enseMsgRegSvr in AMessages) then
    Msg := Msg + '/S ';

  Result:= ExecShellEx(exe, Msg + Fn, '', SW_SHOW, True);
end;

function RegisterNSE(const AFileName: string; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
begin
  Result:= RegUnregNSE(AFileName, True, AMessages);
end;

function UnregisterNSE(const AFileName: string; AMessages: TEasyNSERegMessages = [enseMsgShowErrors]): boolean;
begin
  Result:= RegUnregNSE(AFileName, False, AMessages);
end;


function WideExpandEnviromentString(EnviromentString: string): string;
var
  Length: Integer;
begin
  Result := EnviromentString;
  Length := ExpandEnvironmentStrings(PWideChar( EnviromentString), nil, 0);
  if Length > 0 then
  begin
    SetLength(Result, Length - 1); // Includes the null
    ExpandEnvironmentStrings( PWideChar( EnviromentString), PWideChar( @Result[1]), Length);
  end
end;

function WideExpandEnviromentStringForUser(EnviromentString: string): string;
var
  Token: THandle;
begin
  Result := EnviromentString;
  if OpenProcessToken(GetCurrentProcess, TOKEN_IMPERSONATE or TOKEN_QUERY, Token) then
  begin
    SetLength(Result, 256);
    ExpandEnvironmentStringsForUserW_MP(Token, PWideChar( EnviromentString), PWideChar( @Result[1]), 256);
    SetLength(Result, lstrlenW(PWideChar( Result)));
    CloseHandle(Token)
  end
end;

function WideExtractFileName(Path: string; BaseNameOnly: Boolean = False): string;
var
  i: Integer;
  Found: Boolean;
begin
  Found := False;
  Result := ExtractFileName(Path);
  if BaseNameOnly then
  begin
    i := Length(Result);
    while (i > 0) and not Found do
    begin
      if Result[i] = '.' then
        Found := True
      else
        Dec(i);
    end;
    if Found then
    begin
      Result[i] := #0;
      Result := String(PWideChar( Result));
    end;
  end;
end;

function WideExtractFileDrive(Path: string; DriveLetterOnly: Boolean = False): string;
begin
  Result := ExtractFileDrive(Path);
  if DriveLetterOnly and (Length(Result) > 0) then
    Result := Result[1];
end;

function WideExtractFileExt(Path: string; StripExtPeriod: Boolean = False): string;
var
  Head: PWideChar;
begin
  Result := ExtractFileExt(Path);
  if StripExtPeriod and (Length(Result) > 0) then
  begin
    Head := PWideChar( Result);
    Inc(Head);
    Result := string(Head);
  end;
end;

procedure FixFormFont(AFont: TFont);
var
  LogFont: TLogFont;
begin
  if SystemParametersInfo(SPI_GETICONTITLELOGFONT, SizeOf(LogFont), @LogFont, 0) then
    AFont.Handle := CreateFontIndirect(LogFont)
  else
    AFont.Handle := GetStockObject(DEFAULT_GUI_FONT);
end;

procedure FillGradient(X1, Y1, X2, Y2: integer; fStartColor, fStopColor: TColor;
  StartPoint, EndPoint: integer; fDrawCanvas: TCanvas);
// X1, Y1, X2, Y2: TopLeft and BottomRight coordinates of fill area...
//fStartColor: color to begin the gradient fill with
//FStopColor: color to end the gradient fill with
//StartPoint: the first point between X1 and X2 to draw (useful for faster updating of
// specific areas instead of redrawing the entire gradient fill area such as progress bars)
//EndPoint: the last point between X1 and X2 to draw
//fDrawCanvas: the canvas to draw the gradient on
var
  y: integer;
  tmpColor: TColor;

begin
  fStartColor := ColorToRGB(fStartColor);
  fStopColor := ColorToRGB(fStopColor);
  tmpColor := fDrawCanvas.Pen.Color;
  for y := Y1 to Y2 do begin
    fDrawCanvas.MoveTo(X1, y);
    if (EndPoint > 0) and (y <= EndPoint) and (y >= StartPoint) then begin
      fDrawCanvas.Pen.Color := RGB(Round(GetRValue(fStartColor) + (((GetRValue(fStopColor) - GetRValue(fStartColor)) / (Y2 - Y1)) *Abs(y - Y1))),
          Round(GetGValue(fStartColor) + (((GetGValue(fStopColor) - GetGValue(fStartColor)) / (Y2 - Y1)) * Abs(y - Y1))),
          Round(GetBValue(fStartColor) + (((GetBValue(fStopColor) - GetBValue(fStartColor)) / (y2 - Y1)) * Abs(y - Y1))));
      fDrawCanvas.LineTo(X2, y);
    end;
  end;
  fDrawCanvas.Brush.Color := tmpColor;
end;

procedure ResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight: Integer);
{$if CompilerVersion >= 21}
var
  Factory: IWICImagingFactory;
  Scaler: IWICBitmapScaler;
  Source : TWICImage;
begin
  Bitmap.AlphaFormat := afDefined;
  Source := TWICImage.Create;
  try
    Source.Assign(Bitmap);
    Factory := TWICImage.ImagingFactory;
    Factory.CreateBitmapScaler(Scaler);
    try
      Scaler.Initialize(Source.Handle, NewWidth, NewHeight,
        WICBitmapInterpolationModeHighQualityCubic);
      Source.Handle := IWICBitmap(Scaler);
    finally
      Scaler := nil;
      Factory := nil;
    end;
    Bitmap.Assign(Source);
  finally
    Source.Free;
  end;
{$else}
var
  B: TBitmap;
begin
  B := TBitmap.Create;
  try
    B.SetSize(NewWidth, NewHeight);
    SetStretchBltMode(B.Canvas.Handle, STRETCH_HALFTONE);
    B.Canvas.StretchDraw(Rect(0, 0, NewWidth, NewHeight), Bitmap);
    Bitmap.SetSize(NewWidth, NewHeight);
    Bitmap.Canvas.Draw(0, 0, B);
  finally
    B.Free;
  end;
{$ifend}
end;

function ScaleImageList(Source: TCustomImageList; M, D: Integer): TCustomImageList;
const
  ANDbits: array[0..2*16-1] of  Byte = ($FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF,
                                        $FF,$FF);
  XORbits: array[0..2*16-1] of  Byte = ($00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00,
                                        $00,$00);
var
  I: integer;
  Icon: HIcon;
begin
  Result := TCustomImageList.CreateSize(MulDiv(Source.Width, M, D), MulDiv(Source.Height, M, D));

  if M = D then
  begin
    Result.Assign(Source);
    Exit;
  end;

  Result.ColorDepth := cd32Bit;
  Result.DrawingStyle := Source.DrawingStyle;
  Result.BkColor := Source.BkColor;
  Result.BlendColor := Source.BlendColor;
  for I := 0 to Source.Count-1 do
  begin
    Icon := ImageList_GetIcon(Source.Handle, I, LR_DEFAULTCOLOR);
    if Icon = 0 then
    begin
      Icon := CreateIcon(hInstance,16,16,1,1,@ANDbits,@XORbits);
    end;
    ImageList_AddIcon(Result.Handle, Icon);
    DestroyIcon(Icon);
  end;
end;

function EqualWndMethod(A, B: TWndMethod): Boolean;
begin
  Result := (TMethod(A).Code = TMethod(B).Code) and
            (TMethod(A).Data = TMethod(B).Data)
end;

function FlipReverseCopyRect(const Flip, Reverse: Boolean; const Bitmap: TBitmap): TBitmap;
var
  Bottom, Left, Right, Top:  integer;
begin
  Result := TBitmap.Create;
  Result.Width := Bitmap.Width;
  Result.Height := Bitmap.Height;
  Result.PixelFormat := Bitmap.PixelFormat;

  // Flip Top to Bottom
  if Flip then
  begin
    // Unclear why extra "-1" is needed here.
    Top  := Bitmap.Height-1;
    Bottom := -1
  end
  else begin
    Top  := 0;
    Bottom := Bitmap.Height
  end;

  // Reverse Left to Right
  if Reverse then
  begin
    // Unclear why extra "-1" is needed here.
    Left := Bitmap.Width-1;
    Right := -1;
  end
  else begin
    Left := 0;
    Right := Bitmap.Width;
  end;

  Result.Canvas.CopyRect(Rect(Left,Top, Right,Bottom),
                         Bitmap.Canvas,
                         Rect(0,0, Bitmap.Width,Bitmap.Height));
end;

procedure FlipReverseCopyRect(const Flip, Reverse: Boolean; R: TRect; const Canvas: TCanvas); overload;
var
  Bottom, Left, Right, Top:  integer;
begin
  // Flip Top to Bottom
  if Flip then
  begin
    // Unclear why extra "-1" is needed here.
    Top  := RectHeight(R)-1;
    Bottom := -1
  end
  else begin
    Top  := 0;
    Bottom := RectHeight(R)
  end;

  // Reverse Left to Right
  if Reverse then
  begin
    // Unclear why extra "-1" is needed here.
    Left := RectWidth(R)-1;
    Right := -1;
  end
  else begin
    Left := 0;
    Right := RectWidth(R);
  end;

  Canvas.CopyRect(Rect(Left, Top, Right, Bottom),
                         Canvas,
                         Rect(0,0, RectWidth(R), RectHeight(R)));
end;

function IsMappedDrivePath(const Path: string): Boolean;
var
  WS: string;
begin
  WS := Path;
  SetLength(WS, 3);
  Result := GetDriveType(PWideChar(WS)) = DRIVE_REMOTE
end;

{ Searchs through the passed menu looking for an item identifer that is not   }
{ currently being used.                                                       }

function FindUniqueMenuID(AMenu: HMenu): Cardinal;


    function RunMenu(AMenu: HMenu; var ID: Cardinal): Boolean;
    var
      MenuInfoW: TMenuItemInfoW;
      i, ItemCount: Integer;
      Reset, IsDuplicate: Boolean;
    begin
      Reset := False;
      IsDuplicate := False;
      ItemCount := GetMenuItemCount(AMenu);
      i := 0;
      while (i < ItemCount) and not IsDuplicate do
      begin
        FillChar(MenuInfoW, SizeOf(MenuInfoW), #0);
        MenuInfoW.cbSize := SizeOf(MenuInfoW);
        MenuInfoW.fMask := MIIM_SUBMENU or MIIM_ID;
        GetMenuItemInfoW(AMenu, i, True, MenuInfoW);
        if MenuInfoW.hSubMenu <> 0 then
          Reset := RunMenu(MenuInfoW.hSubMenu, ID);
        IsDuplicate := MenuInfoW.wID = ID;
        Inc(i);
      end;
      Result := IsDuplicate and not Reset
    end;

begin
  Result := UniqueMenuIDSeed;
  while RunMenu(AMenu, Result) do
    Inc(Result);
  UniqueMenuIDSeed := Result;
  Inc(UniqueMenuIDSeed);
  if UniqueMenuIDSeed > 32000 then
    UniqueMenuIDSeed := 1000;
end;

function AddContextMenuItem(Menu: HMenu; ACaption: string; Index: Integer;
  MenuID: UINT = $FFFF; hSubMenu: UINT = 0; Enabled: Boolean = True;
  Checked: Boolean = False; Default: Boolean = False): Integer;
//
// Pass '-' for a separator
//      -1 to add to the end
//      if MenuID = -1 then the function will create a unique ID
//      if hSubMenu > 0 then the item contain sub-items
// Returns ID of new Item
//
var
  InfoW: TMenuItemInfoW;
begin
  FillChar(InfoW, SizeOf(InfoW), #0);
  InfoW.cbSize := SizeOf(InfoW);
  InfoW.fMask := MIIM_TYPE or MIIM_ID or MIIM_STATE;

  if Enabled or (ACaption = '-') then
    InfoW.fState := InfoW.fState or MFS_ENABLED
  else
    InfoW.fState := InfoW.fState or MFS_DISABLED;

  if Checked and (ACaption <> '-') then
    InfoW.fState := InfoW.fState or MFS_CHECKED;
  if Default and (ACaption <> '-') then
    InfoW.fState := InfoW.fState or MFS_DEFAULT;

  if ACaption = '-' then
    InfoW.fType := MFT_SEPARATOR
  else begin
    InfoW.fType := MFT_STRING;
    if hSubMenu > 0 then
    begin
      InfoW.fMask := InfoW.fMask or MIIM_SUBMENU;
      InfoW.hSubMenu := hSubMenu
    end
  end;
  InfoW.dwTypeData := PWideChar(ACaption);
  InfoW.cch := Length(ACaption);

  if InfoW.fType = MFT_STRING then
  begin
    if MenuID = $FFFF then
      InfoW.wID := FindUniqueMenuID(Menu)
    else
    if InfoW.fMask and MIIM_SUBMENU <> 0 then
      InfoW.wID := $FFFF // Sub-Item Parents don't get an unique ID
    else
      InfoW.wID := MenuID;
  end else
    InfoW.wID := $FFFF; // Separators don't get an unique ID

  Result := InfoW.wID;
  if Index < 0 then
    InsertMenuItem(Menu, GetMenuItemCount(Menu), True, InfoW)
  else
    InsertMenuItem(Menu, Index, True, InfoW);  // Inserts by Position
end;

procedure ShadowBlendBits(Bits: TBitmap; BackGndColor: TColor);
begin
  if Assigned(Bits) and (Bits.PixelFormat = pf32Bit) then
  begin
    AlphaBlend(Bits.Canvas.Handle, Bits.Canvas.Handle,
          Rect(0, 0, Bits.Width,Bits.Height), Point(0, 0),
          cbmConstantAlphaAndColor, 0, ColorToRGB(BackGndColor));
    ConvertBitmapEx(Bits, Bits, BackGndColor)
  end
end;


procedure SumFolder(FolderPath: string; Recurse: Boolean; var Size: Int64);
{ Returns the size of all files within the passed folder, including all         }
{ sub-folders. This is recurcive don't initialize Size to 0 in the function!    }
var
  InfoW: TWin32FindDataW;
  FHandle: THandle;
begin
  FHandle := FindFirstFile(PWideChar( FolderPath + '\*.*'), InfoW);
  if FHandle <> INVALID_HANDLE_VALUE then
  try
    if Recurse and (InfoW.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY <> 0) then
    begin
      if (lstrcmpi(InfoW.cFileName, '.') <> 0) and (lstrcmpi(InfoW.cFileName, '..') <> 0) and
        (InfoW.dwFileAttributes and FILE_ATTRIBUTE_REPARSE_POINT = 0) then
        SumFolder(FolderPath + '\' + InfoW.cFileName, Recurse, Size)
    end else
      Size := Size + Int64(InfoW.nFileSizeHigh) * MAXDWORD + Int64(InfoW.nFileSizeLow);
    while FindNextFile(FHandle, InfoW) and not SumFolderAbort do
    begin
      if Recurse and (InfoW.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY <> 0) then
      begin
        if (lstrcmpi(InfoW.cFileName, '.') <> 0) and (lstrcmpi(InfoW.cFileName, '..') <> 0) and
          (InfoW.dwFileAttributes and FILE_ATTRIBUTE_REPARSE_POINT = 0) then
         SumFolder(FolderPath + '\' + InfoW.cFileName, Recurse, Size)
      end else
       Size := Size + Int64(InfoW.nFileSizeHigh) * MAXDWORD + Int64(InfoW.nFileSizeLow);
    end;
  finally
    Windows.FindClose(FHandle)
  end
end;

function InternalTextExtentW(Text: PWideChar; DC: HDC): TSize;
begin
  GetTextExtentPoint32W(DC, PWideChar(Text), Length(Text), Result);
end;

function WideValidateDelimitedExtList(DelimitedText: string;
  Prefix: TValidateDelimiterWildCardSet; Delimiter: TValidateDelimiterExt): string;
//
// Forces the passes string of delimited extensions to to following format:
//
//  *.ext;*.bak;*.txt  when WideValidateDelimitedExtList = True
//  ext;bak;txt        when WideValidateDelimitedExtList = False
//
//  The passed string can contain the following delimiters
//
//  ";", ":", ",", "|" that will be converted to ";"
//
//  Spaces will be stripped and the characters will be lowercased, empty strings
// will be removed. Also the following ill-formed strings will be corrected:
//
//  "*txt", ".txt"
//
//  to "*.txt"
//
var
  i: Integer;
  TestStr: string;
  Extensions: TStringList;
begin
  Result := '';
  Extensions := TStringList.Create;
  try
    Extensions.Sorted := False;
    case Delimiter of
      vdeColon:
        begin
          DelimitedText := StringReplace(DelimitedText, ',', ':', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, ';', ':', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, '|', ':', [rfReplaceAll]);
          Extensions.Delimiter := ':';
        end;
      vdeSemiColon:
        begin
          DelimitedText := StringReplace(DelimitedText, ',', ';', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, ':', ';', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, '|', ';', [rfReplaceAll]);
          Extensions.Delimiter := ';';
        end;
      vdeComma:
        begin
          DelimitedText := StringReplace(DelimitedText, ':', ',', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, ';', ',', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, '|', ',', [rfReplaceAll]);
          Extensions.Delimiter := ',';
        end;
      vdePipe:
        begin
          DelimitedText := StringReplace(DelimitedText, ':', '|', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, ';', '|', [rfReplaceAll]);
          DelimitedText := StringReplace(DelimitedText, ',', '|', [rfReplaceAll]);
          Extensions.Delimiter := '|';
        end
      end;


    Extensions.DelimitedText := DelimitedText;

    for i := Extensions.Count - 1 downto 0 do
    begin
      TestStr := SysUtils.AnsiLowerCase(Trim(Extensions[i]));     // Strip off white space
      if Length(TestStr) < 1 then
        Extensions.Delete(i)               // Remove if it has no characters
      else begin

        // Strip both prefixes if avaialable
        if (TestStr[1] = '*') or (TestStr[1] = '.')  then
        begin
          TestStr[1] := ' ';
          if (TestStr[2] = '*') or (TestStr[2] = '.')  then
            TestStr[2] := ' ';
          TestStr := Trim(TestStr);
        end;

        if [vdwcAsterisk, vdwcPeriod] * Prefix = [vdwcAsterisk, vdwcPeriod] then
          Extensions[i] := '*.' + TestStr
        else
        if vdwcAsterisk in Prefix then
          Extensions[i] := '*' + TestStr
        else
        if vdwcPeriod in Prefix then
          Extensions[i] := '.' + TestStr
        else
          Extensions[i] := TestStr;

      end
    end;
    Extensions.Sorted := True;
    Result := Extensions.DelimitedText;
  finally
    Extensions.Free;
  end
end;

function LibList: TList;
begin
  if not Assigned(FLibList) then
    FLibList := TList.Create;
  Result := FLibList
end;

function CommonLoadLibrary(LibraryName: string): THandle;
var
  i: Integer;
  Found: Boolean;
  LibRec: PLibRec;
begin
  Result := 0;
  Found := False;
  i := 0;
  while (i < LibList.Count) and not Found do
  begin
    LibRec := PLibRec(LibList[i]);
    if lstrcmpi(PWideChar(LibRec.LibraryName), PWideChar(LibraryName)) = 0 then
    begin
      Inc(LibRec.ReferenceCount);
      Result := LibRec.Handle;
      Found := True
    end;
    Inc(i)
  end;
  if not Found then
  begin
    New(LibRec);
    LibRec.Handle := LoadLibrary(PWideChar(LibraryName));
    if LibRec.Handle <> 0 then
    begin
      LibRec.LibraryName := LibraryName;
      LibRec.ReferenceCount := 1;
      LibList.Add(LibRec);
      Result := LibRec.Handle
    end else
    Dispose(LibRec)
  end
end;

function CommonUnloadLibrary(LibraryName: string): Boolean;
var
  i: Integer;
  LibRec: PLibRec;
begin
  Result := False;
  i := 0;
  while (i < LibList.Count) and not Result do
  begin
    LibRec := PLibRec(LibList[i]);
    if lstrcmpi(PWideChar(LibRec.LibraryName), PWideChar(LibraryName)) = 0 then
    begin
      Dec(LibRec.ReferenceCount);
      FreeLibrary(LibRec.Handle);
      if LibRec.ReferenceCount = 0 then
      begin
        LibList.Delete(i);
        Dispose(LibRec);
      end;
      Result := True
    end;
    Inc(i)
  end;
  if LibList.Count = 0 then
    FreeAndNil(FLibList)
end;

procedure CommonUnloadAllLibraries;
var
  LibIndex: Integer;
  LibRec: PLibRec;
begin
  while (LibList.Count - 1) >= 0 do
  begin
    LibRec := PLibRec(LibList[LibList.Count - 1]);
    for LibIndex := 0 to LibRec.ReferenceCount - 1 do
      FreeLibrary(LibRec.Handle);
    LibList.Delete(LibList.Count - 1);
    Dispose(LibRec);
  end;
  FreeAndNil(FLibList);
end;

procedure DrawWindowButton(Canvas: TCanvas; Pos: TPoint; Size: Integer; ButtonType: TDrawWindowButtonType);
begin
  MarlettFont.Size := Size;
  Canvas.Brush.Style := bsClear;
  Canvas.Font.Assign(MarlettFont);
  case ButtonType of
    dwbtMinimize: Canvas.TextOut(Pos.X, Pos.Y, '0');
    dwbtMaximize: Canvas.TextOut(Pos.X, Pos.Y, '1');
    dwbtRestore:  Canvas.TextOut(Pos.X, Pos.Y, '2');
    dwbtClose:    Canvas.TextOut(Pos.X, Pos.Y, 'r');
  end;

end;

procedure DrawRadioButton(Canvas: TCanvas; Pos: TPoint; Size: Integer; clBackground, clHotBkGnd,
  clLeftOuter, clRightOuter, clLeftInner, clRightInner: TColor; Checked, Enabled, Hot: Boolean);
begin
  MarlettFont.Size := Size;
  Canvas.Brush.Style := bsClear;
  Canvas.Font.Assign(MarlettFont);

    // Draw the background
  if Hot then
    Canvas.Font.Color := clHotBkGnd
  else
    Canvas.Font.Color := clBackground;
  Canvas.TextOut(Pos.X, Pos.Y, 'n');

  Canvas.Brush.Style := bsClear;
  // Draw the Outer Circle
  Canvas.Font.Color := clLeftOuter;
  Canvas.TextOut(Pos.X, Pos.Y, 'j');
  Canvas.Font.Color := clRightOuter;
  Canvas.TextOut(Pos.X, Pos.Y, 'k');
  // Draw the Inner Circle
  Canvas.Font.Color := clLeftInner;
  Canvas.TextOut(Pos.X, Pos.Y, 'l');
  Canvas.Font.Color := clRightInner;
  Canvas.TextOut(Pos.X, Pos.Y, 'm');
  if Checked then
  begin
    if Enabled then
      Canvas.Font.Color := clBlack
    else
      Canvas.Font.Color := clBtnShadow;
    Canvas.TextOut(Pos.X, Pos.Y, 'i');
  end
end;

procedure DrawCheckBox(Canvas: TCanvas; Pos: TPoint; Size: Integer; clBackground, clHotBkGnd,
  clLeftOuter, clRightOuter, clLeftInner, clRightInner: TColor; Checked, Enabled, Hot: Boolean);
begin
  MarlettFont.Size := Size;
  Canvas.Brush.Style := bsSolid;
  Canvas.Font.Assign(MarlettFont);

  // Draw the background
  if Hot then
    Canvas.Font.Color := clHotBkGnd
  else
    Canvas.Font.Color := clBackground;
  Canvas.TextOut(Pos.X, Pos.Y, Char($67));

  Canvas.Brush.Style := bsClear;
  // Draw the Outer Frame
  Canvas.Font.Color := clLeftOuter;
  Canvas.TextOut(Pos.X, Pos.Y, Char($63));
  Canvas.Font.Color := clRightOuter;
  Canvas.TextOut(Pos.X, Pos.Y, Char($64));
  // Draw the Inner Frame
  Canvas.Font.Color := clLeftInner;
  Canvas.TextOut(Pos.X, Pos.Y, Char($65));
  Canvas.Font.Color := clRightInner;
  Canvas.TextOut(Pos.X, Pos.Y, Char($66));
  if Checked then
  begin
    if Enabled then
      Canvas.Font.Color := clBlack
    else
      Canvas.Font.Color := clBtnShadow;
    Canvas.TextOut(Pos.X, Pos.Y, Char($62));
  end
end;

function CheckBounds(Size: Integer; Character: Char = Char($67)): TRect;
var
  Canvas: TCanvas;
begin
  Result := Rect(0, 0, 0, 0);
  Canvas := TCanvas.Create;
  try
    Canvas.Handle := GetDC(0);
    Canvas.Font.Name := 'Marlett';
    Canvas.Font.Size := Size;
    // Use the background for the Checkbox for the size, the Radio will be this
    // size or a bit smaller
    Result.Right := Canvas.TextWidth(Character);
    Result.Bottom := Canvas.TextHeight(Character);
  finally
    if Assigned(Canvas) then
    begin
      ReleaseDC(0, Canvas.Handle);
      Canvas.Handle := 0
    end;
    Canvas.Free
  end;
end;


function HasMMX: Boolean;

// Helper method to determine whether the current processor supports MMX.

{$ifdef CPUX64}
begin
  // We use SSE2 in the "MMX-functions"
  Result := True;
end;
{$else}
asm
        PUSH    EBX
        XOR     EAX, EAX     // Result := False
        PUSHFD               // determine if the processor supports the CPUID command
        POP     EDX
        MOV     ECX, EDX
        XOR     EDX, $200000
        PUSH    EDX
        POPFD
        PUSHFD
        POP     EDX
        XOR     ECX, EDX
        JZ      @1           // no CPUID support so we can't even get to the feature information
        PUSH    EDX
        POPFD

        MOV     EAX, 1
        DW      $A20F        // CPUID, EAX contains now version info and EDX feature information
        MOV     EBX, EAX     // free EAX to get the result value
        XOR     EAX, EAX     // Result := False
        CMP     EBX, $50
        JB      @1           // if processor family is < 5 then it is not a Pentium class processor
        TEST    EDX, $800000
        JZ      @1           // if the MMX bit is not set then we don't have MMX
        INC     EAX          // Result := True
@1:
        POP     EBX
end;
{$endif CPUX64}

procedure AlphaBlendLineConstant(Source, Destination: Pointer; Count: Integer; ConstantAlpha, Bias: Integer);

// Blends a line of Count pixels from Source to Destination using a constant alpha value.
// The layout of a pixel must be BGRA where A is ignored (but is calculated as the other components).
// ConstantAlpha must be in the range 0..255 where 0 means totally transparent (destination pixel only)
// and 255 totally opaque (source pixel only).
// Bias is an additional value which gets added to every component and must be in the range -128..127
//
{$ifdef CPUX64}
// RCX contains Source
// RDX contains Destination
// R8D contains Count
// R9D contains ConstantAlpha
// Bias is on the stack

asm
        //.NOFRAME

        // Load XMM3 with the constant alpha value (replicate it for every component).
        // Expand it to word size.
        MOVD        XMM3, R9D  // ConstantAlpha
        PUNPCKLWD   XMM3, XMM3
        PUNPCKLDQ   XMM3, XMM3

        // Load XMM5 with the bias value.
        MOVD        XMM5, [Bias]
        PUNPCKLWD   XMM5, XMM5
        PUNPCKLDQ   XMM5, XMM5

        // Load XMM4 with 128 to allow for saturated biasing.
        MOV         R10D, 128
        MOVD        XMM4, R10D
        PUNPCKLWD   XMM4, XMM4
        PUNPCKLDQ   XMM4, XMM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        MOVD        XMM1, DWORD PTR [RCX]   // data is unaligned
        MOVD        XMM2, DWORD PTR [RDX]   // data is unaligned
        PXOR        XMM0, XMM0    // clear source pixel register for unpacking
        PUNPCKLBW   XMM0, XMM1{[RCX]}    // unpack source pixel byte values into words
        PSRLW       XMM0, 8       // move higher bytes to lower bytes
        PXOR        XMM1, XMM1    // clear target pixel register for unpacking
        PUNPCKLBW   XMM1, XMM2{[RDX]}    // unpack target pixel byte values into words
        MOVQ        XMM2, XMM1    // make a copy of the shifted values, we need them again
        PSRLW       XMM1, 8       // move higher bytes to lower bytes

        // calculation is: target = (alpha * (source - target) + 256 * target) / 256
        PSUBW       XMM0, XMM1    // source - target
        PMULLW      XMM0, XMM3    // alpha * (source - target)
        PADDW       XMM0, XMM2    // add target (in shifted form)
        PSRLW       XMM0, 8       // divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        PSUBW     XMM0, XMM4
        PADDSW    XMM0, XMM5
        PADDW     XMM0, XMM4
        PACKUSWB  XMM0, XMM0      // convert words to bytes with saturation
        MOVD      DWORD PTR [RDX], XMM0     // store the result
@3:
        ADD       RCX, 4
        ADD       RDX, 4
        DEC       R8D
        JNZ       @1
end;
{$else}
// EAX contains Source
// EDX contains Destination
// ECX contains Count
// ConstantAlpha and Bias are on the stack

asm
        PUSH    ESI                    // save used registers
        PUSH    EDI

        MOV     ESI, EAX               // ESI becomes the actual source pointer
        MOV     EDI, EDX               // EDI becomes the actual target pointer

        // Load MM6 with the constant alpha value (replicate it for every component).
        // Expand it to word size.
        MOV     EAX, [ConstantAlpha]
        DB      $0F, $6E, $F0          /// MOVD      MM6, EAX
        DB      $0F, $61, $F6          /// PUNPCKLWD MM6, MM6
        DB      $0F, $62, $F6          /// PUNPCKLDQ MM6, MM6

        // Load MM5 with the bias value.
        MOV     EAX, [Bias]
        DB      $0F, $6E, $E8          /// MOVD      MM5, EAX
        DB      $0F, $61, $ED          /// PUNPCKLWD MM5, MM5
        DB      $0F, $62, $ED          /// PUNPCKLDQ MM5, MM5

        // Load MM4 with 128 to allow for saturated biasing.
        MOV     EAX, 128
        DB      $0F, $6E, $E0          /// MOVD      MM4, EAX
        DB      $0F, $61, $E4          /// PUNPCKLWD MM4, MM4
        DB      $0F, $62, $E4          /// PUNPCKLDQ MM4, MM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        DB      $0F, $EF, $C0          /// PXOR      MM0, MM0,   clear source pixel register for unpacking
        DB      $0F, $60, $06          /// PUNPCKLBW MM0, [ESI], unpack source pixel byte values into words
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     move higher bytes to lower bytes
        DB      $0F, $EF, $C9          /// PXOR      MM1, MM1,   clear target pixel register for unpacking
        DB      $0F, $60, $0F          /// PUNPCKLBW MM1, [EDI], unpack target pixel byte values into words
        DB      $0F, $6F, $D1          /// MOVQ      MM2, MM1,   make a copy of the shifted values, we need them again
        DB      $0F, $71, $D1, $08     /// PSRLW     MM1, 8,     move higher bytes to lower bytes

        // calculation is: target = (alpha * (source - target) + 256 * target) / 256
        DB      $0F, $F9, $C1          /// PSUBW     MM0, MM1,   source - target
        DB      $0F, $D5, $C6          /// PMULLW    MM0, MM6,   alpha * (source - target)
        DB      $0F, $FD, $C2          /// PADDW     MM0, MM2,   add target (in shifted form)
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        DB      $0F, $F9, $C4          /// PSUBW     MM0, MM4
        DB      $0F, $ED, $C5          /// PADDSW    MM0, MM5
        DB      $0F, $FD, $C4          /// PADDW     MM0, MM4
        DB      $0F, $67, $C0          /// PACKUSWB  MM0, MM0,   convert words to bytes with saturation
        DB      $0F, $7E, $07          /// MOVD      [EDI], MM0, store the result
@3:
        ADD     ESI, 4
        ADD     EDI, 4
        DEC     ECX
        JNZ     @1
        POP     EDI
        POP     ESI
end;
{$endif CPUX64}

//----------------------------------------------------------------------------------------------------------------------

procedure AlphaBlendLinePerPixel(Source, Destination: Pointer; Count, Bias: Integer);

// Blends a line of Count pixels from Source to Destination using the alpha value of the source pixels.
// The layout of a pixel must be BGRA.
// Bias is an additional value which gets added to every component and must be in the range -128..127
//
{$ifdef CPUX64}
// RCX contains Source
// RDX contains Destination
// R8D contains Count
// R9D contains Bias

asm
        //.NOFRAME

        // Load XMM5 with the bias value.
        MOVD        XMM5, R9D   // Bias
        PUNPCKLWD   XMM5, XMM5
        PUNPCKLDQ   XMM5, XMM5

        // Load XMM4 with 128 to allow for saturated biasing.
        MOV         R10D, 128
        MOVD        XMM4, R10D
        PUNPCKLWD   XMM4, XMM4
        PUNPCKLDQ   XMM4, XMM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        MOVD        XMM1, DWORD PTR [RCX]   // data is unaligned
        MOVD        XMM2, DWORD PTR [RDX]   // data is unaligned
        PXOR        XMM0, XMM0    // clear source pixel register for unpacking
        PUNPCKLBW   XMM0, XMM1{[RCX]}    // unpack source pixel byte values into words
        PSRLW       XMM0, 8       // move higher bytes to lower bytes
        PXOR        XMM1, XMM1    // clear target pixel register for unpacking
        PUNPCKLBW   XMM1, XMM2{[RDX]}    // unpack target pixel byte values into words
        MOVQ        XMM2, XMM1    // make a copy of the shifted values, we need them again
        PSRLW       XMM1, 8       // move higher bytes to lower bytes

        // Load XMM3 with the source alpha value (replicate it for every component).
        // Expand it to word size.
        MOVQ        XMM3, XMM0
        PUNPCKHWD   XMM3, XMM3
        PUNPCKHDQ   XMM3, XMM3

        // calculation is: target = (alpha * (source - target) + 256 * target) / 256
        PSUBW       XMM0, XMM1    // source - target
        PMULLW      XMM0, XMM3    // alpha * (source - target)
        PADDW       XMM0, XMM2    // add target (in shifted form)
        PSRLW       XMM0, 8       // divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        PSUBW       XMM0, XMM4
        PADDSW      XMM0, XMM5
        PADDW       XMM0, XMM4
        PACKUSWB    XMM0, XMM0    // convert words to bytes with saturation
        MOVD        DWORD PTR [RDX], XMM0   // store the result
@3:
        ADD         RCX, 4
        ADD         RDX, 4
        DEC         R8D
        JNZ         @1
end;
{$else}
// EAX contains Source
// EDX contains Destination
// ECX contains Count
// Bias is on the stack

asm
        PUSH    ESI                    // save used registers
        PUSH    EDI

        MOV     ESI, EAX               // ESI becomes the actual source pointer
        MOV     EDI, EDX               // EDI becomes the actual target pointer

        // Load MM5 with the bias value.
        MOV     EAX, [Bias]
        DB      $0F, $6E, $E8          /// MOVD      MM5, EAX
        DB      $0F, $61, $ED          /// PUNPCKLWD MM5, MM5
        DB      $0F, $62, $ED          /// PUNPCKLDQ MM5, MM5

        // Load MM4 with 128 to allow for saturated biasing.
        MOV     EAX, 128
        DB      $0F, $6E, $E0          /// MOVD      MM4, EAX
        DB      $0F, $61, $E4          /// PUNPCKLWD MM4, MM4
        DB      $0F, $62, $E4          /// PUNPCKLDQ MM4, MM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        DB      $0F, $EF, $C0          /// PXOR      MM0, MM0,   clear source pixel register for unpacking
        DB      $0F, $60, $06          /// PUNPCKLBW MM0, [ESI], unpack source pixel byte values into words
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     move higher bytes to lower bytes
        DB      $0F, $EF, $C9          /// PXOR      MM1, MM1,   clear target pixel register for unpacking
        DB      $0F, $60, $0F          /// PUNPCKLBW MM1, [EDI], unpack target pixel byte values into words
        DB      $0F, $6F, $D1          /// MOVQ      MM2, MM1,   make a copy of the shifted values, we need them again
        DB      $0F, $71, $D1, $08     /// PSRLW     MM1, 8,     move higher bytes to lower bytes

        // Load MM6 with the source alpha value (replicate it for every component).
        // Expand it to word size.
        DB      $0F, $6F, $F0          /// MOVQ MM6, MM0
        DB      $0F, $69, $F6          /// PUNPCKHWD MM6, MM6
        DB      $0F, $6A, $F6          /// PUNPCKHDQ MM6, MM6

        // calculation is: target = (alpha * (source - target) + 256 * target) / 256
        DB      $0F, $F9, $C1          /// PSUBW     MM0, MM1,   source - target
        DB      $0F, $D5, $C6          /// PMULLW    MM0, MM6,   alpha * (source - target)
        DB      $0F, $FD, $C2          /// PADDW     MM0, MM2,   add target (in shifted form)
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        DB      $0F, $F9, $C4          /// PSUBW     MM0, MM4
        DB      $0F, $ED, $C5          /// PADDSW    MM0, MM5
        DB      $0F, $FD, $C4          /// PADDW     MM0, MM4
        DB      $0F, $67, $C0          /// PACKUSWB  MM0, MM0,   convert words to bytes with saturation
        DB      $0F, $7E, $07          /// MOVD      [EDI], MM0, store the result
@3:
        ADD     ESI, 4
        ADD     EDI, 4
        DEC     ECX
        JNZ     @1
        POP     EDI
        POP     ESI
end;
{$endif CPUX64}

//----------------------------------------------------------------------------------------------------------------------

procedure AlphaBlendLineMaster(Source, Destination: Pointer; Count: Integer; ConstantAlpha, Bias: Integer);

// Blends a line of Count pixels from Source to Destination using the source pixel and a constant alpha value.
// The layout of a pixel must be BGRA.
// ConstantAlpha must be in the range 0..255.
// Bias is an additional value which gets added to every component and must be in the range -128..127
//
{$ifdef CPUX64}
// RCX contains Source
// RDX contains Destination
// R8D contains Count
// R9D contains ConstantAlpha
// Bias is on the stack

asm
        .SAVENV XMM6

        // Load XMM3 with the constant alpha value (replicate it for every component).
        // Expand it to word size.
        MOVD        XMM3, R9D    // ConstantAlpha
        PUNPCKLWD   XMM3, XMM3
        PUNPCKLDQ   XMM3, XMM3

        // Load XMM5 with the bias value.
        MOV         R10D, [Bias]
        MOVD        XMM5, R10D
        PUNPCKLWD   XMM5, XMM5
        PUNPCKLDQ   XMM5, XMM5

        // Load XMM4 with 128 to allow for saturated biasing.
        MOV         R10D, 128
        MOVD        XMM4, R10D
        PUNPCKLWD   XMM4, XMM4
        PUNPCKLDQ   XMM4, XMM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        MOVD        XMM1, DWORD PTR [RCX]   // data is unaligned
        MOVD        XMM2, DWORD PTR [RDX]   // data is unaligned
        PXOR        XMM0, XMM0    // clear source pixel register for unpacking
        PUNPCKLBW   XMM0, XMM1{[RCX]}     // unpack source pixel byte values into words
        PSRLW       XMM0, 8       // move higher bytes to lower bytes
        PXOR        XMM1, XMM1    // clear target pixel register for unpacking
        PUNPCKLBW   XMM1, XMM2{[RCX]}     // unpack target pixel byte values into words
        MOVQ        XMM2, XMM1    // make a copy of the shifted values, we need them again
        PSRLW       XMM1, 8       // move higher bytes to lower bytes

        // Load XMM6 with the source alpha value (replicate it for every component).
        // Expand it to word size.
        MOVQ        XMM6, XMM0
        PUNPCKHWD   XMM6, XMM6
        PUNPCKHDQ   XMM6, XMM6
        PMULLW      XMM6, XMM3    // source alpha * master alpha
        PSRLW       XMM6, 8       // divide by 256

        // calculation is: target = (alpha * master alpha * (source - target) + 256 * target) / 256
        PSUBW       XMM0, XMM1    // source - target
        PMULLW      XMM0, XMM6    // alpha * (source - target)
        PADDW       XMM0, XMM2    // add target (in shifted form)
        PSRLW       XMM0, 8       // divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        PSUBW       XMM0, XMM4
        PADDSW      XMM0, XMM5
        PADDW       XMM0, XMM4
        PACKUSWB    XMM0, XMM0    // convert words to bytes with saturation
        MOVD        DWORD PTR [RDX], XMM0   // store the result
@3:
        ADD         RCX, 4
        ADD         RDX, 4
        DEC         R8D
        JNZ         @1
end;
{$else}
// EAX contains Source
// EDX contains Destination
// ECX contains Count
// ConstantAlpha and Bias are on the stack

asm
        PUSH    ESI                    // save used registers
        PUSH    EDI

        MOV     ESI, EAX               // ESI becomes the actual source pointer
        MOV     EDI, EDX               // EDI becomes the actual target pointer

        // Load MM6 with the constant alpha value (replicate it for every component).
        // Expand it to word size.
        MOV     EAX, [ConstantAlpha]
        DB      $0F, $6E, $F0          /// MOVD      MM6, EAX
        DB      $0F, $61, $F6          /// PUNPCKLWD MM6, MM6
        DB      $0F, $62, $F6          /// PUNPCKLDQ MM6, MM6

        // Load MM5 with the bias value.
        MOV     EAX, [Bias]
        DB      $0F, $6E, $E8          /// MOVD      MM5, EAX
        DB      $0F, $61, $ED          /// PUNPCKLWD MM5, MM5
        DB      $0F, $62, $ED          /// PUNPCKLDQ MM5, MM5

        // Load MM4 with 128 to allow for saturated biasing.
        MOV     EAX, 128
        DB      $0F, $6E, $E0          /// MOVD      MM4, EAX
        DB      $0F, $61, $E4          /// PUNPCKLWD MM4, MM4
        DB      $0F, $62, $E4          /// PUNPCKLDQ MM4, MM4

@1:     // The pixel loop calculates an entire pixel in one run.
        // Note: The pixel byte values are expanded into the higher bytes of a word due
        //       to the way unpacking works. We compensate for this with an extra shift.
        DB      $0F, $EF, $C0          /// PXOR      MM0, MM0,   clear source pixel register for unpacking
        DB      $0F, $60, $06          /// PUNPCKLBW MM0, [ESI], unpack source pixel byte values into words
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     move higher bytes to lower bytes
        DB      $0F, $EF, $C9          /// PXOR      MM1, MM1,   clear target pixel register for unpacking
        DB      $0F, $60, $0F          /// PUNPCKLBW MM1, [EDI], unpack target pixel byte values into words
        DB      $0F, $6F, $D1          /// MOVQ      MM2, MM1,   make a copy of the shifted values, we need them again
        DB      $0F, $71, $D1, $08     /// PSRLW     MM1, 8,     move higher bytes to lower bytes

        // Load MM7 with the source alpha value (replicate it for every component).
        // Expand it to word size.
        DB      $0F, $6F, $F8          /// MOVQ      MM7, MM0
        DB      $0F, $69, $FF          /// PUNPCKHWD MM7, MM7
        DB      $0F, $6A, $FF          /// PUNPCKHDQ MM7, MM7
        DB      $0F, $D5, $FE          /// PMULLW    MM7, MM6,   source alpha * master alpha
        DB      $0F, $71, $D7, $08     /// PSRLW     MM7, 8,     divide by 256

        // calculation is: target = (alpha * master alpha * (source - target) + 256 * target) / 256
        DB      $0F, $F9, $C1          /// PSUBW     MM0, MM1,   source - target
        DB      $0F, $D5, $C7          /// PMULLW    MM0, MM7,   alpha * (source - target)
        DB      $0F, $FD, $C2          /// PADDW     MM0, MM2,   add target (in shifted form)
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8,     divide by 256

        // Bias is accounted for by conversion of range 0..255 to -128..127,
        // doing a saturated add and convert back to 0..255.
        DB      $0F, $F9, $C4          /// PSUBW     MM0, MM4
        DB      $0F, $ED, $C5          /// PADDSW    MM0, MM5
        DB      $0F, $FD, $C4          /// PADDW     MM0, MM4
        DB      $0F, $67, $C0          /// PACKUSWB  MM0, MM0,   convert words to bytes with saturation
        DB      $0F, $7E, $07          /// MOVD      [EDI], MM0, store the result
@3:
        ADD     ESI, 4
        ADD     EDI, 4
        DEC     ECX
        JNZ     @1
        POP     EDI
        POP     ESI
end;
{$endif CPUX64}

//----------------------------------------------------------------------------------------------------------------------

procedure AlphaBlendLineMasterAndColor(Destination: Pointer; Count: Integer; ConstantAlpha, Color: Integer);

// Blends a line of Count pixels in Destination against the given color using a constant alpha value.
// The layout of a pixel must be BGRA and Color must be rrggbb00 (as stored by a COLORREF).
// ConstantAlpha must be in the range 0..255.
//
{$ifdef CPUX64}
// RCX contains Destination
// EDX contains Count
// R8D contains ConstantAlpha
// R9D contains Color

asm
        //.NOFRAME

        // The used formula is: target = (alpha * color + (256 - alpha) * target) / 256.
        // alpha * color (factor 1) and 256 - alpha (factor 2) are constant values which can be calculated in advance.
        // The remaining calculation is therefore: target = (F1 + F2 * target) / 256

        // Load XMM3 with the constant alpha value (replicate it for every component).
        // Expand it to word size. (Every calculation here works on word sized operands.)
        MOVD        XMM3, R8D   // ConstantAlpha
        PUNPCKLWD   XMM3, XMM3
        PUNPCKLDQ   XMM3, XMM3

        // Calculate factor 2.
        MOV         R10D, $100
        MOVD        XMM2, R10D
        PUNPCKLWD   XMM2, XMM2
        PUNPCKLDQ   XMM2, XMM2
        PSUBW       XMM2, XMM3             // XMM2 contains now: 255 - alpha = F2

        // Now calculate factor 1. Alpha is still in XMM3, but the r and b components of Color must be swapped.
        BSWAP       R9D  // Color
        ROR         R9D, 8
        MOVD        XMM1, R9D              // Load the color and convert to word sized values.
        PXOR        XMM4, XMM4
        PUNPCKLBW   XMM1, XMM4
        PMULLW      XMM1, XMM3             // XMM1 contains now: color * alpha = F1

@1:     // The pixel loop calculates an entire pixel in one run.
        MOVD        XMM0, DWORD PTR [RCX]
        PUNPCKLBW   XMM0, XMM4

        PMULLW      XMM0, XMM2             // calculate F1 + F2 * target
        PADDW       XMM0, XMM1
        PSRLW       XMM0, 8                // divide by 256

        PACKUSWB    XMM0, XMM0             // convert words to bytes with saturation
        MOVD        DWORD PTR [RCX], XMM0            // store the result

        ADD         RCX, 4
        DEC         EDX
        JNZ         @1
end;
{$else}
// EAX contains Destination
// EDX contains Count
// ECX contains ConstantAlpha
// Color is passed on the stack

asm
        // The used formula is: target = (alpha * color + (256 - alpha) * target) / 256.
        // alpha * color (factor 1) and 256 - alpha (factor 2) are constant values which can be calculated in advance.
        // The remaining calculation is therefore: target = (F1 + F2 * target) / 256

        // Load MM3 with the constant alpha value (replicate it for every component).
        // Expand it to word size. (Every calculation here works on word sized operands.)
        DB      $0F, $6E, $D9          /// MOVD      MM3, ECX
        DB      $0F, $61, $DB          /// PUNPCKLWD MM3, MM3
        DB      $0F, $62, $DB          /// PUNPCKLDQ MM3, MM3

        // Calculate factor 2.
        MOV     ECX, $100
        DB      $0F, $6E, $D1          /// MOVD      MM2, ECX
        DB      $0F, $61, $D2          /// PUNPCKLWD MM2, MM2
        DB      $0F, $62, $D2          /// PUNPCKLDQ MM2, MM2
        DB      $0F, $F9, $D3          /// PSUBW     MM2, MM3             // MM2 contains now: 255 - alpha = F2

        // Now calculate factor 1. Alpha is still in MM3, but the r and b components of Color must be swapped.
        MOV     ECX, [Color]
        BSWAP   ECX
        ROR     ECX, 8
        DB      $0F, $6E, $C9          /// MOVD      MM1, ECX             // Load the color and convert to word sized values.
        DB      $0F, $EF, $E4          /// PXOR      MM4, MM4
        DB      $0F, $60, $CC          /// PUNPCKLBW MM1, MM4
        DB      $0F, $D5, $CB          /// PMULLW    MM1, MM3             // MM1 contains now: color * alpha = F1

@1:     // The pixel loop calculates an entire pixel in one run.
        DB      $0F, $6E, $00          /// MOVD      MM0, [EAX]
        DB      $0F, $60, $C4          /// PUNPCKLBW MM0, MM4

        DB      $0F, $D5, $C2          /// PMULLW    MM0, MM2             // calculate F1 + F2 * target
        DB      $0F, $FD, $C1          /// PADDW     MM0, MM1
        DB      $0F, $71, $D0, $08     /// PSRLW     MM0, 8               // divide by 256

        DB      $0F, $67, $C0          /// PACKUSWB  MM0, MM0             // convert words to bytes with saturation
        DB      $0F, $7E, $00          /// MOVD      [EAX], MM0           // store the result

        ADD     EAX, 4
        DEC     EDX
        JNZ     @1
end;
{$endif CPUX64}

//----------------------------------------------------------------------------------------------------------------------

procedure EMMS;

// Reset MMX state to use the FPU for other tasks again.

{$ifdef CPUX64}
  inline;
begin
end;
{$else}
asm
        DB      $0F, $77               /// EMMS
end;
{$endif CPUX64}

function GetBitmapBitsFromDeviceContext(DC: HDC; var Width, Height: Integer): Pointer;
// Helper function used to retrieve the bitmap selected into the given device context. If there is a bitmap then
// the function will return a pointer to its bits otherwise nil is returned.
// Additionally the dimensions of the bitmap are returned.
var
  Bitmap: HBITMAP;
  DIB: TDIBSection;
begin
  Result := nil;
  Width := 0;
  Height := 0;

  Bitmap := GetCurrentObject(DC, OBJ_BITMAP);
  if Bitmap <> 0 then
  begin
    if GetObject(Bitmap, SizeOf(DIB), @DIB) = SizeOf(DIB) then
    begin
      Assert(DIB.dsBm.bmPlanes * DIB.dsBm.bmBitsPixel = 32, 'Alpha blending error: bitmap must use 32 bpp.');
      Result := DIB.dsBm.bmBits;
      Width := DIB.dsBmih.biWidth;
      Height := DIB.dsBmih.biHeight;
    end;
  end;
  Assert(Result <> nil, 'Alpha blending DC error: no bitmap available.');
end;

function CalculateScanline(Bits: Pointer; Width, Height, Row: Integer): Pointer;

// Helper function to calculate the start address for the given row.

begin
  if Height > 0 then  // bottom-up DIB
    Row := Height - Row - 1;
  // Return DWORD aligned address of the requested scanline.
  Result := PAnsiChar(Bits) + Row * ((Width * 32 + 31) and not 31) div 8;
end;

procedure ConvertBitmapEx(Image32: TBitmap; var OutImage: TBitmap; const BackGndColor: TColor);
var
  I, N: Integer;
  LongColor: DWORD;
  SourceRed, SourceGreen, SourceBlue, BkGndRed, BkGndGreen, BkGndBlue, RedTarget, GreenTarget, BlueTarget, Alpha: Byte;
  Target, Mask: TBitmap;
  LineDeltaImage32, PixelDeltaImage32, LineDeltaTarget, PixelDeltaTarget, LineDeltaMask, PixelDeltaMask: Integer;
  PLineImage32, PLineTarget, PLineMask, PPixelImage32, PPixelTarget, PPixelMask: PByte;
begin
  // Algorithm only works for bitmaps with a height > 1 pixel, should not be a limitation
  // as it would then be a line!
  if (Image32.PixelFormat = pf32Bit) and (Image32.Height > 1) and UsesAlphaChannel(Image32) then
  begin
    Target := TBitmap.Create;
    Mask := TBitmap.Create;

    try
      Target.PixelFormat := pf32Bit;
      Target.Width := Image32.Width;
      Target.Height := Image32.Height;
      Target.Assign(Image32);

      Mask.PixelFormat := pf32Bit;
      Mask.Width := Image32.Width;
      Mask.Height := Image32.Height;
      Mask.Canvas.Brush.Color := BackGndColor;
      Mask.Canvas.FillRect(Mask.Canvas.ClipRect);

      LineDeltaImage32 := DWORD( Image32.ScanLine[1]) - DWORD( Image32.ScanLine[0]);
      LineDeltaTarget := DWORD( Target.ScanLine[1]) - DWORD( Target.ScanLine[0]);
      LineDeltaMask := DWORD( Mask.ScanLine[1]) - DWORD( Mask.ScanLine[0]);

      PixelDeltaImage32 := SizeOf(TRGBQuad);
      PixelDeltaTarget := SizeOf(TRGBQuad);
      PixelDeltaMask := SizeOf(TRGBQuad);

      PLineImage32 := Image32.ScanLine[0];
      PLineTarget := Target.ScanLine[0];
      PLineMask := Mask.ScanLine[0];

      for I := 0 to Image32.Height - 1 do
      begin
        PPixelImage32 := PLineImage32;
        PPixelTarget := PLineTarget;
        PPixelMask := PLineMask;

        for N := 0 to Image32.Width - 1 do
        begin
          // Source GetColorValues ; Profiled = ~24-30% of time
          LongColor := PDWORD( PPixelImage32)^;
          SourceBlue := LongColor and $000000FF;
          SourceGreen := (LongColor and $0000FF00) shr 8;
          SourceRed := (LongColor and $00FF0000) shr 16;
          Alpha := (LongColor and $FF000000) shr 24;

          // Mask GetColorValues ; Profiled = ~24-30% of time
          LongColor := PDWORD( PPixelMask)^;
          BkGndBlue := LongColor and $000000FF;
          BkGndGreen := (LongColor and $0000FF00) shr 8;
          BkGndRed := (LongColor and $00FF0000) shr 16;

          // displayColor = sourceColor�alpha / 256 + backgroundColor�(256 � alpha) / 256
          // Profiled = ~15-24% of time
          RedTarget := SourceRed*Alpha shr 8 + BkGndRed*(255-Alpha) shr 8;
          GreenTarget := SourceGreen*Alpha shr 8 + BkGndGreen*(255-Alpha) shr 8;
          BlueTarget := SourceBlue*Alpha shr 8 + BkGndBlue*(255-Alpha) shr 8;

      //    RedTarget := BkGndRed*Alpha shr 8 + SourceRed*(255-Alpha) shr 8;
      //    GreenTarget := BkGndGreen*Alpha shr 8 + SourceGreen*(255-Alpha) shr 8;
      //    BlueTarget := BkGndBlue*Alpha shr 8 + SourceBlue*(255-Alpha) shr 8;


          // Create the RGB DWORD color ; Profiled = ~8%-9%% of time
          // Mask out all but the alpha channel then build the backwards stored RGB preserving the alpha channel bits
          PDWORD(PPixelTarget)^ := ((BlueTarget) or (GreenTarget shl 8)or (RedTarget shl 16));

          Inc(PPixelImage32, PixelDeltaImage32);
          Inc(PPixelTarget, PixelDeltaTarget);
          Inc(PPixelMask, PixelDeltaMask);
        end;
        Inc(PLineImage32, LineDeltaImage32);
        Inc(PLineTarget, LineDeltaTarget);
        Inc(PLineMask, LineDeltaMask);
      end;
      OutImage.Assign(Target);
    finally
      FreeAndNil(Target);
      FreeAndNil(Mask);
    end;
  end else
   OutImage.Assign(Image32)
end;

procedure AlphaBlend(Source, Destination: HDC; R: TRect; Target: TPoint; Mode: TCommonBlendMode; ConstantAlpha, Bias: Integer);

// NOTE:::::::::::::
//    AlphaBlend does not respect any clipping in the DC!!!!!!!
//

// Optimized alpha blend procedure using MMX instructions to perform as quick as possible.
// For this procedure to work properly it is important that both source and target bitmap use the 32 bit color format.
// R describes the source rectangle to work on.
// Target is the place (upper left corner) in the target bitmap where to blend to. Note that source width + X offset
// must be less or equal to the target width. Similar for the height.
// If Mode is bmConstantAlpha then the blend operation uses the given ConstantAlpha value for all pixels.
// If Mode is bmPerPixelAlpha then each pixel is blended using its individual alpha value (the alpha value of the source).
// If Mode is bmMasterAlpha then each pixel is blended using its individual alpha value multiplied by ConstantAlpha.
// If Mode is bmConstantAlphaAndColor then each destination pixel is blended using ConstantAlpha but also a constant
// color which will be obtained from Bias. In this case no offset value is added, otherwise Bias is used as offset.
// Blending of a color into target only (bmConstantAlphaAndColor) ignores Source (the DC) and Target (the position).
// CAUTION: This procedure does not check whether MMX instructions are actually available! Call it only if MMX is really
//          usable.

var
  Y: Integer;
  SourceRun,
  TargetRun: PByte;

  SourceBits,
  DestBits: Pointer;
  SourceWidth,
  SourceHeight,
  DestWidth,
  DestHeight: Integer;

begin
  if not IsRectEmpty(R) then
  begin
    // Note: it is tempting to optimize the special cases for constant alpha 0 and 255 by just ignoring soure
    //       (alpha = 0) or simply do a blit (alpha = 255). But this does not take the bias into account.
    case Mode of
      cbmConstantAlpha:
        begin
          // Get a pointer to the bitmap bits for the source and target device contexts.
          // Note: this supposes that both contexts do actually have bitmaps assigned!
          SourceBits := GetBitmapBitsFromDeviceContext(Source, SourceWidth, SourceHeight);
          DestBits := GetBitmapBitsFromDeviceContext(Destination, DestWidth, DestHeight);
          if Assigned(SourceBits) and Assigned(DestBits) then
          begin
            for Y := 0 to R.Bottom - R.Top - 1 do
            begin
              SourceRun := CalculateScanline(SourceBits, SourceWidth, SourceHeight, Y + R.Top);
              Inc(SourceRun, 4 * R.Left);
              TargetRun := CalculateScanline(DestBits, DestWidth, DestHeight, Y + Target.Y);
              Inc(TargetRun, 4 * Target.X);
              AlphaBlendLineConstant(SourceRun, TargetRun, R.Right - R.Left, ConstantAlpha, Bias);
            end;
          end;
          EMMS;
        end;
      cbmPerPixelAlpha:
        begin
          SourceBits := GetBitmapBitsFromDeviceContext(Source, SourceWidth, SourceHeight);
          DestBits := GetBitmapBitsFromDeviceContext(Destination, DestWidth, DestHeight);
          if Assigned(SourceBits) and Assigned(DestBits) then
          begin
            for Y := 0 to R.Bottom - R.Top - 1 do
            begin
              SourceRun := CalculateScanline(SourceBits, SourceWidth, SourceHeight, Y + R.Top);
              Inc(SourceRun, 4 * R.Left);
              TargetRun := CalculateScanline(DestBits, DestWidth, DestHeight, Y + Target.Y);
              Inc(TargetRun, 4 * Target.X);
              AlphaBlendLinePerPixel(SourceRun, TargetRun, R.Right - R.Left, Bias);
            end;
          end;
          EMMS;
        end;
      cbmMasterAlpha:
        begin
          SourceBits := GetBitmapBitsFromDeviceContext(Source, SourceWidth, SourceHeight);
          DestBits := GetBitmapBitsFromDeviceContext(Destination, DestWidth, DestHeight);
          if Assigned(SourceBits) and Assigned(DestBits) then
          begin
            for Y := 0 to R.Bottom - R.Top - 1 do
            begin
              SourceRun := CalculateScanline(SourceBits, SourceWidth, SourceHeight, Y + R.Top);
              Inc(SourceRun, 4 * Target.X);
              TargetRun := CalculateScanline(DestBits, DestWidth, DestHeight, Y + Target.Y);
              AlphaBlendLineMaster(SourceRun, TargetRun, R.Right - R.Left, ConstantAlpha, Bias);
            end;
          end;
          EMMS;
        end;
      cbmConstantAlphaAndColor:
        begin
          // Source is ignore since there is a constant color value.
          DestBits := GetBitmapBitsFromDeviceContext(Destination, DestWidth, DestHeight);
          if Assigned(DestBits) then
          begin
            for Y := 0 to R.Bottom - R.Top - 1 do
            begin
              TargetRun := CalculateScanline(DestBits, DestWidth, DestHeight, Y + R.Top);
              Inc(TargetRun, 4 * R.Left);
              AlphaBlendLineMasterAndColor(TargetRun, R.Right - R.Left, ConstantAlpha, Bias);
            end;
          end;
          EMMS;
        end;
    end;
  end;
end;

function DrawTextWEx(DC: HDC; Text: string; var lpRect: TRect;
  Flags: TCommonDrawTextWFlags; MaxLineCount: Integer): Integer;
// Creates and extented version of DrawTextW that works in Win9x as well as
// NT.  If MaxLineCount is -1 then the line count will depend on the Text.  All
// lines that are extracted from the text are drawn or calcuated in the rectangle
//
// The result is the number of lines actually drawn, note if the CalcRect flags are
//   used the result will be the number of lines that would be drawn
var
  TextMetrics: TTextMetric;
  Size: TSize;
  TextPosX, TextPosY, i, NewLineTop: Integer;
  TextOutFlags: Longword;
  LineRect, OldlpRect: TRect;
  Buffer: TCommonWideCharArray;
  BufferIndex: PWideChar;
  ShortText: string;
  VOffset, SplitCount: Integer;
begin
  OldlpRect := lpRect;
  GetTextMetrics(DC, TextMetrics);

  TextOutFlags := 0;
  if dtRTLReading in Flags then
    TextOutFlags := TextOutFlags or ETO_RTLREADING;
  if not (dtNoClip in Flags) then
    TextOutFlags := TextOutFlags or ETO_CLIPPED;

  if dtSingleLine in Flags then
  begin
    Result := 1;  // Easy one!

    // Set up the LineRect in the Vertical Direction
    // Default to the top
    LineRect := Rect(lpRect.Left, lpRect.Top, lpRect.Right, lpRect.Top + TextMetrics.tmHeight);
    if dtVCenter in Flags then
      OffsetRect(LineRect, 0, (RectHeight(lpRect) - RectHeight(LineRect)) div 2)
    else
    if dtBottom in Flags then
      OffsetRect(LineRect, 0, RectHeight(lpRect) - RectHeight(LineRect));
    TextPosX := LineRect.Left;
    TextPosY := LineRect.Top;

    if dtEndEllipsis in Flags then
      Text := ShortenTextW(DC, Text, RectWidth(LineRect));

    GetTextExtentPoint32W(DC, PWideChar(Text), Length(Text), Size);
    if Flags * [dtCenter, dtRight] <> [] then
    begin
      if dtCenter in Flags then
        TextPosX := TextPosX + (RectWidth(LineRect) - Size.cx) div 2
      else
        TextPosX := LineRect.Right - Size.cx
    end;

 {   if dtCenter in Flags then
      SetTextAlign(DC, TA_CENTER)
    else
    if dtLeft in Flags then
      SetTextAlign(DC, TA_LEFT)
    else
    if dtRight in Flags then
      SetTextAlign(DC, TA_RIGHT);   }

    // See if the caller wants to only calculate the rectangle for the text
    if dtCalcRect in Flags then
    begin
      // Assume that the text will fit in the calulated line/rect
      lpRect.Left := TextPosX;
      lpRect.Top := TextPosY;
      lpRect.Bottom := LineRect.Bottom;
      lpRect.Right := lpRect.Left + Size.cx;

      // If it does not then we have to do some adjusting
      if Size.cx > RectWidth(OldlpRect) then
      begin
        if dtCalcRectAlign in Flags then
        begin
          lpRect.Left := OldlpRect.Left;
          lpRect.Right := OldlpRect.Right;
        end;
        if dtCalcRectAdjR in Flags then
          lpRect.Right := lpRect.Left + Size.cx;
      end
    end else
      ExtTextOutW(DC, TextPosX, TextPosY, TextOutFlags, @LineRect, PWideChar(Text), Length(Text), nil);

  end else
  begin
    // It is multi-line
    SplitCount := SplitTextW(DC, Text, lpRect.Right-lpRect.Left, Buffer, MaxLineCount);
    i := 0;
    if Length(Buffer) > 0 then
    begin
      // We call ourselves recursivly one line at a time to draw the multi line text
      Include(Flags, dtSingleLine);
      BufferIndex := @Buffer[0];

      // Calculate where the center of the text block is with respect to the
      // rectangle
  {    if dtVCenter in Flags then
      begin
        if (SplitCount > MaxLineCount) and (MaxLineCount > -1) then
          VOffset := (RectHeight(OldlpRect) - (TextMetrics.tmHeight * MaxLineCount)) div 2
        else
          VOffset := (RectHeight(OldlpRect) - (TextMetrics.tmHeight * SplitCount)) div 2;
        if VOffset < 0 then
          VOffset := 0;
      end else
      if dtBottom in Flags then
      begin
        VOffset := (RectHeight(OldlpRect) - (TextMetrics.tmHeight * MaxLineCount))
      end else
        VOffset := 0; }

      // Fix for multitext vertical alignment from Solerman Kaplon 11.9.04
      if (dtVCenter in Flags) or (dtBottom in Flags) then
      begin
        if (SplitCount > MaxLineCount) and (MaxLineCount > -1) then
        VOffset := (RectHeight(OldlpRect) - (TextMetrics.tmHeight * MaxLineCount))
      else
      VOffset := (RectHeight(OldlpRect) - (TextMetrics.tmHeight * SplitCount));
      if VOffset < 0 then
        VOffset := 0
      else
      if dtVCenter in Flags then
        VOffset := VOffset shr 1;
      end else
        VOffset := 0;
      while ((i < MaxLineCount) or (MaxLineCount < 0)) and (BufferIndex^ <> WideNull) do
      begin
        // Calculate where the top of a single line of text starts
        NewLineTop := OldlpRect.Top + (i * TextMetrics.tmHeight) + VOffset;
        LineRect := Rect(OldlpRect.Left, NewLineTop, OldlpRect.Right, NewLineTop + TextMetrics.tmHeight);
        if (dtEndEllipsis in Flags) {and not(dtCalcRect in Flags)} then
        begin
          ShortText := ShortenTextW(DC, string(BufferIndex), RectWidth(OldlpRect));
          DrawTextWEx(DC, ShortText, LineRect, Flags, MaxLineCount);
        end else
          DrawTextWEx(DC, WideString(BufferIndex), LineRect, Flags, MaxLineCount);

        if dtCalcRect in Flags then
        begin
          if i = 0 then
            lpRect := LineRect
          else
            UnionRect(lpRect, lpRect, LineRect);
        end;
        Inc(BufferIndex, lStrLenW(BufferIndex) + 1);
        Inc(i)
      end;

      if (SplitCount = 0) and (dtCalcRect in Flags) then
      begin
        if dtCalcRectAdjR in Flags then
        begin
          lpRect.Right := lpRect.Left;
          lpRect.Bottom := lpRect.Top + TextMetrics.tmHeight
        end else
        begin
          lpRect.Bottom := lpRect.Top;
          lpREct.Right := lpRect.Left + TextMetrics.tmAveCharWidth
        end
      end;

      if SplitCount > MaxLineCount then
        Result := MaxLineCount
      else
        Result := SplitCount
    end else
      Result := 0;
  end
end;

procedure CreateProcessMP(ExeFile, Parameters, InitalDir: string);
var
  pi: TProcessInformation;
  siW: TStartupInfo;
  wA, wB, wC: PWideChar;
begin
  FillChar(pi, SizeOf(pi), #0);
  FillChar(siW, SizeOf(siW), #0);
  wA := nil;
  wB := nil;
  wC := nil;
  if ExeFile <> '' then
    wA := PWideChar(ExeFile);
  if Parameters <> '' then
    wB := PWideChar(Parameters);
  if InitalDir <> '' then
    wC := PWideChar(InitalDir);

  CreateProcess(
                wA, // path to the executable file:
                wB,
                nil,
                nil,
                False,
                NORMAL_PRIORITY_CLASS,
                nil,
                wC,
                siW,
                pi );
  if pi.hProcess <> 0 then
    CloseHandle(pi.hProcess);
  if pi.hThread <> 0 then
    CloseHandle(pi.hThread)
end;


function DiffRectHorz(Rect1, Rect2: TRect): TRect;
// Returns the "difference" rectangle of the passed rects in the Horz direction.
// Assumes that one corner is common between the two rects
begin
  Rect1 := ProperRect(Rect1);
  Rect2 := ProperRect(Rect2);
  // Make sure we contain every thing horizontally
  Result.Left := Min(Rect1.Left, Rect2.Left);
  Result.Right := Max(Rect1.Right, Rect1.Right);
  // Now find the difference rect height
  if Rect1.Top = Rect2.Top then
  begin
    // The tops are equal so it must be the bottom that contains the difference
    Result.Bottom := Max(Rect1.Bottom, Rect2.Bottom);
    Result.Top := Min(Rect1.Bottom, Rect2.Bottom);
  end else
  begin
   // The bottoms are equal so it must be the tops that contains the difference
    Result.Bottom := Max(Rect1.Top, Rect2.Top);
    Result.Top := Min(Rect1.Top, Rect2.Top);
  end
end;

function DiffRectVert(Rect1, Rect2: TRect): TRect;
// Returns the "difference" rectangle of the passed rects in the Vert direction.
// Assumes that one corner is common between the two rects
begin
  Rect1 := ProperRect(Rect1);
  Rect2 := ProperRect(Rect2);
  // Make sure we contain every thing vertically
  Result.Top := Min(Rect1.Top, Rect2.Bottom);
  Result.Bottom := Max(Rect1.Top, Rect1.Bottom);
  // Now find the difference rect width
  if Rect1.Left = Rect2.Left then
  begin
    // The tops are equal so it must be the bottom that contains the difference
    Result.Right := Max(Rect1.Right, Rect2.Right);
    Result.Left := Min(Rect1.Right, Rect2.Right);
  end else
  begin
   // The bottoms are equal so it must be the tops that contains the difference
    Result.Right := Max(Rect1.Left, Rect2.Left);
    Result.Left := Min(Rect1.Left, Rect2.Left);
  end
end;

function AbsRect(ARect: TRect): TRect;
// Makes all coodinates positive
begin
  Result := ARect;
  if Result.Left < 0 then
    Result.Left := 0;
  if Result.Top < 0 then
    Result.Top := 0;
  if Result.Right < 0 then
    Result.Right := 0;
  if Result.Bottom < 0 then
    Result.Bottom := 0;
end;

function CenterRectInRect(OuterRect, InnerRect: TRect): TRect;
begin
  if RectWidth(InnerRect) > RectWidth(OuterRect) then
  begin
    // If the inner rect is wider than the result is the outer rect x bounds
    Result.Left := OuterRect.Left;
    Result.Right := OuterRect.Right;
  end else
  begin
    // If not then center the inner rectangle in the outer in the x direction
    Result.Left := OuterRect.Left;
    Result.Right := Result.Left + RectWidth(InnerRect);
    OffsetRect(Result, (RectWidth(OuterRect) - RectWidth(InnerRect)) div 2, 0);
  end;
  if RectHeight(InnerRect) > RectHeight(OuterRect) then
  begin
    // If the inner rect is wider than the result is the outer rect y bounds
    Result.Top := OuterRect.Top;
    Result.Bottom := OuterRect.Bottom;
  end else
  begin
    // If not then center the inner rectangle in the outer in the y direction
    Result.Top := OuterRect.Top;
    Result.Bottom := Result.Top + RectHeight(InnerRect);
    OffsetRect(Result, 0, (RectHeight(OuterRect) - RectHeight(InnerRect)) div 2);
  end;
end;

function CenterRectHorz(OuterRect, InnerRect: TRect): TRect;
begin
  if RectWidth(InnerRect) > RectWidth(OuterRect) then
  begin
    // If the inner rect is wider than the result is the outer rect x bounds
    Result.Left := OuterRect.Left;
    Result.Right := OuterRect.Right;
  end else
  begin
    // If not then center the inner rectangle in the outer in the x direction
    Result.Left := OuterRect.Left;
    Result.Right := Result.Left + RectWidth(InnerRect);
    OffsetRect(Result, (RectWidth(OuterRect) - RectWidth(InnerRect)) div 2, 0);
  end;
end;

function CenterRectVert(OuterRect, InnerRect: TRect): TRect;
begin
if RectHeight(InnerRect) > RectHeight(OuterRect) then
  begin
    // If the inner rect is wider than the result is the outer rect y bounds
    Result.Top := OuterRect.Top;
    Result.Bottom := OuterRect.Bottom;
  end else
  begin
    // If not then center the inner rectangle in the outer in the y direction
    Result.Top := OuterRect.Top;
    Result.Bottom := Result.Top + RectHeight(InnerRect);
    OffsetRect(Result, 0, (RectHeight(OuterRect) - RectHeight(InnerRect)) div 2);
  end;
end;

function CommonSupports(const Instance: IUnknown; const IID: TGUID; out Intf): Boolean; overload;
begin
  Result := (Instance <> nil) and (Instance.QueryInterface(IID, Intf) = S_OK);
end;

function CommonSupports(const Instance: TObject; const IID: TGUID; out Intf): Boolean; overload;
var
  LUnknown: IUnknown;
begin
  Result := (Instance <> nil) and
            (Instance.GetInterface(IID, Intf) or
            (Supports(LUnknown, IID, Intf) and Instance.GetInterface(IUnknown, LUnknown)));
end;

function CommonSupports(const Instance: IUnknown; const IID: TGUID): Boolean; overload;
var
  Temp: IUnknown;
begin
  Result := Supports(Instance, IID, Temp);
end;

function CommonSupports(const Instance: TObject; const IID: TGUID): Boolean; overload;
 var
  Temp: IUnknown;
begin
  Result := Supports(Instance, IID, Temp);
end;

procedure MinMax(var A, B: Integer);
// Makes sure that A < B
var
  Temp: Integer;
begin
  if A > B then
  begin
    Temp := A;
    A := B;
    B := Temp
  end
end;

function IsRectProper(Rect: TRect): Boolean;
begin
  Result := (Rect.Right >= Rect.Left) and (Rect.Bottom >= Rect.Top)
end;

function AsyncLeftButtonDown: Boolean;
begin
  Result := GetKeyState(VK_LBUTTON) and $8000 <> 0
end;

function AsyncRightButtonDown: Boolean;
begin
  Result := GetKeyState(VK_RBUTTON) and $8000 <> 0
end;

function AsyncMiddleButtonDown: Boolean;
begin
  Result := GetKeyState(VK_MBUTTON) and $8000 <> 0
end;

function AsyncControlDown: Boolean;
begin
  Result := GetKeyState(VK_CONTROL) and $8000 <> 0
end;

function AsyncAltDown: Boolean;
begin
  Result := (GetKeyState(VK_RMENU) and $8000 <> 0) or(GetKeyState(VK_LMENU) and $8000 <> 0)
end;

function AsyncShiftDown: Boolean;
begin
  Result := GetKeyState(VK_SHIFT) and $8000 <> 0
end;

function CalcuateFolderSize(FolderPath: string; Recurse: Boolean): Int64;

// Recursivly gets the size of the folder and subfolders
var
  FreeSpaceAvailable, TotalSpace: Int64;
begin
  Result := 0;
  if Recurse and WideIsDrive(FolderPath) then
  begin
    if SysUtils.GetDiskFreeSpaceEx(PWideChar(FolderPath), FreeSpaceAvailable, TotalSpace, nil) then
    Result := TotalSpace - FreeSpaceAvailable
  end else
  begin
    SumFolderAbort := False;
    SumFolder(FolderPath, Recurse, Result);
  end
end;

function GetMyDocumentsVirtualFolder: PItemIDList;

const
  MYCOMPUTER_GUID = WideString('::{450d8fba-ad25-11d0-98a8-0800361b1103}');

var
  dwAttributes, pchEaten: ULONG;
  Desktop: IShellFolder;
begin
  Result := nil;
  dwAttributes := 0;
  SHGetDesktopFolder(Desktop);
  pchEaten := Length(MYCOMPUTER_GUID);
  if not Succeeded(Desktop.ParseDisplayName(0, nil,
    PWideChar(MYCOMPUTER_GUID), pchEaten, Result, dwAttributes))
  then
    Result := nil
end;

function WideGetTempDir: string;
var
  BufferW: array[0..MAX_PATH] of Widechar;
begin
  if GetTempPath(MAX_PATH, BufferW) > 0 then
    Result := BufferW;
end;

function WideIncrementalSearch(CompareStr, Mask: string): Integer;
begin
  SetLength(CompareStr, Length(Mask));
  Result := lstrcmpi(PWideChar(Mask), PWideChar(CompareStr));
end;

function WideIsDrive(Drive: string): Boolean;
begin
  if Length(Drive) = 3 then
    Result := (LowerCase(Drive[1]) >= 'a') and (LowerCase(Drive[1]) <= 'z') and (Drive[2] = ':') and (Drive[3] = '\')
  else
  if Length(Drive) = 2 then
    Result := (LowerCase(Drive[1]) >= 'a') and (LowerCase(Drive[1]) <= 'z') and (Drive[2] = ':')
  else
    Result := False
end;

function WideIsFloppy(FileFolder: string): boolean;
begin
  if Length(FileFolder) > 0 then
    Result := WideIsDrive(FileFolder) and
      ((WideChar(FileFolder[1]) = WideChar('A')) or
      (WideChar(FileFolder[1]) = WideChar('a')) or
      (WideChar(FileFolder[1]) = WideChar('B')) or
      (WideChar(FileFolder[1]) = WideChar('b')))
  else
    Result := False
end;


function IsAnyMouseButtonDown: Boolean;
begin
  Result := not(((GetAsyncKeyState(VK_LBUTTON) and $8000) = 0) and
            ((GetAsyncKeyState(VK_RBUTTON) and $8000) = 0) and
            ((GetAsyncKeyState(VK_MBUTTON) and $8000) = 0))
end;

function WideNewFolderName(ParentFolder: string; SuggestedFolderName: string = ''): string;
var
  i: integer;
  TempFoldername: String;
begin
  ParentFolder := WideStripTrailingBackslash(ParentFolder, True); // Strip even if a root folder
  i := 1;
  if SuggestedFolderName = '' then
    Begin
      Result := ParentFolder + '\' + STR_NEWFOLDER;
      TempFoldername := STR_NEWFOLDER;
    end
  else
    Begin
      Result := ParentFolder + '\' + SuggestedFolderName;
      Tempfoldername := SuggestedFolderName;
    End;
  while DirectoryExists(PWideChar(Result)) and (i <= High(WORD)) do
  begin
    Result := ParentFolder + '\' + Tempfoldername + ' (' + IntToStr(i) + ')';
    Inc(i);
  end;
  if i > High(WORD) then
    Result := '';
end;

function WidePathMatchSpec(Path, Mask: string): Boolean;
begin
  if Assigned(PathMatchSpecW_MP) then
    Result := PathMatchSpecW_MP(PWideChar(Path), PWideChar( Mask))
  else
    Result := False
end;

function WidePathMatchSpecExists: Boolean;
begin
  Result := Assigned(PathMatchSpecW_MP)
end;

function WideIsPathDelimiter(const S: string; Index: Integer): Boolean;
begin
  Result := (Index > 0) and (Index <= Length(S)) and (S[Index] = '\');
end;

function IsTextTrueType(DC: HDC): Boolean;
var
    TextMetrics: TTextMetric;
begin
   GetTextMetrics(DC, TextMetrics);
   Result := TextMetrics.tmPitchAndFamily and TMPF_TRUETYPE <>  0
end;

function IsTextTrueType(Canvas: TCanvas): Boolean;
begin
  Result := IsTextTrueType(Canvas.Handle);
end;

function IsUNCPath(const Path: string): Boolean;
begin
  if Length(Path) > 2 then
    Result :=  ((Path[1] = '\') and (Path[2] = '\')) and (DirectoryExists(PWideChar(Path)) or FileExists(Path))
  else
    Result := False
end;

function IsUNCPathSyntax(const Path: string): Boolean;
begin
  if Length(Path) > 2 then
    Result :=  ((Path[1] = '\') and (Path[2] = '\'))
  else
    Result := False
end;

function StrRetToStr(StrRet: TStrRet; APIDL: PItemIDList): string;
{ Extracts the string from the StrRet structure.                                }
var
  P: PAnsiChar;
begin
  case StrRet.uType of
    STRRET_CSTR:
      begin
        P := @StrRet.cStr[0];
        SetString(Result, P, Length(P));
      end;
    STRRET_OFFSET:
      begin
        if Assigned(APIDL) then
        begin
          {$R-}
          P := PAnsiChar(@(APIDL).mkid.abID[StrRet.uOffset - SizeOf(APIDL.mkid.cb)]);
          {$R+}
          SetString(Result, P, Length(P));
        end else
          Result := '';
      end;
    STRRET_WSTR:
    begin
      Result := StrRet.pOleStr;
      if Assigned(StrRet.pOleStr) then
        PIDLMgr.FreeOLEStr(StrRet.pOLEStr);
     end;
  end;
end;

function SystemDirectory: string;
var
  Len: integer;
begin
  Result := '';
  Len := GetSystemDirectory(PWideChar(Result), 0);
  begin
    SetLength(Result, Len - 1);
    GetSystemDirectory(PWideChar(Result), Len);
  end
end;

function SysMenuFont: HFONT;
var
  MetricsW: TNonClientMetricsW;
begin
  FillChar(MetricsW, SizeOf(MetricsW), #0);
  MetricsW.cbSize := SizeOf(MetricsW);
  SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Sizeof(MetricsW), @MetricsW, 0);
  Result := CreateFontIndirect(MetricsW.lfMenuFont);
end;

function SysMenuHeight: Integer;
var
  MetricsW: TNonClientMetricsW;
begin
  FillChar(MetricsW, SizeOf(MetricsW), #0);
  MetricsW.cbSize := SizeOf(MetricsW);
  SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Sizeof(MetricsW), @MetricsW, 0);
  Result := MetricsW.iMenuHeight
end;

function TextExtentW(Text: string; Font: TFont): TSize;
var
  Canvas: TCanvas;
begin
  FillChar(Result, SizeOf(Result), #0);
  if Text <> '' then
  begin
    Canvas := TCanvas.Create;
    try
      Canvas.Handle := GetDC(0);
      Canvas.Lock;
      Canvas.Font.Assign(Font);
      Result := InternalTextExtentW(PWideChar(Text), Canvas.Handle);
    finally
      if Assigned(Canvas) and (Canvas.Handle <> 0) then
        ReleaseDC(0, Canvas.Handle);
      Canvas.Unlock;
      Canvas.Free
    end
  end
end;

function TextExtentW(Text: string; Canvas: TCanvas): TSize;
begin
  FillChar(Result, SizeOf(Result), #0);
  if Assigned(Canvas) and (Text <> '') then
  begin
    Canvas.Lock;
    Result := InternalTextExtentW(PWideChar(Text), Canvas.Handle);
    Canvas.Unlock;
  end;
end;

function TextExtentW(Text: PWideChar; Canvas: TCanvas): TSize;
begin
  FillChar(Result, SizeOf(Result), #0);
  if Assigned(Canvas) and (Assigned(Text)) then
  begin
    Canvas.Lock;
    Result := InternalTextExtentW(Text, Canvas.Handle);
    Canvas.Unlock;
  end;
end;

function TextExtentW(Text: PWideChar; DC: hDC): TSize;
begin
  FillChar(Result, SizeOf(Result), #0);
  if (DC <> 0) and (Assigned(Text)) then
    Result := InternalTextExtentW(Text, DC);
end;

type
  TABCArray = array of TABC;

function TextTrueExtentsW(Text: string; DC: HDC): TSize;
var
  ABC: TABC;
  TextMetrics: TTextMetric;
  i: integer;
begin
   // Get the Height at least
   GetTextExtentPoint32W(DC, PWideChar(Text), Length(Text), Result);

   GetTextMetrics(DC, TextMetrics);
   if TextMetrics.tmPitchAndFamily and TMPF_TRUETYPE <> 0 then
   begin
     Result.cx := 0;
     // Is TrueType
     for i := 1 to Length(Text) do
     begin
       GetCharABCWidths(DC, Ord(Text[i]), Ord(Text[i]), ABC);
       Result.cx := Result.cx + ABC.abcA + integer(ABC.abcB) + ABC.abcC;
     end
   end
end;

function UniqueFileName(const AFilePath: string): string;

{ Creates a unique file name in based on other files in the passed path         }

var
  i: integer;
  WP: PWideChar;
begin
  Result := AFilePath;
  i := 2;
  while FileExists(Result) and (i < 20) do
  begin
    Result := AFilePath;
    WP := SysUtils.StrRScan(PWideChar( Result), '.');
    if Assigned(WP) then
      Insert( ' (' + IntToStr(i) + ')', Result, PWideChar(WP) - PWideChar(Result))
    else begin
      Result := '';
      Break;
    end;
    Inc(i)
  end;
end;

function UniqueDirName(const ADirPath: string): string;
var
  i: integer;
begin
  Result := ADirPath;
  i := 2;
  while DirectoryExists(PWideChar(Result)) and (i < 20) do
  begin
    Result := ADirPath;
    Insert( ' (' + IntToStr(i) + ')', Result, Length(Result));
    Inc(i)
  end;
end;

function WideStripExt(AFile: string): string;
{ Strips the extenstion off a file name }
var
  i: integer;
  Done: Boolean;
begin
  i := Length(AFile);
  Done := False;
  Result := AFile;
  while (i > 0) and not Done do
  begin
    if AFile[i] = '.' then
    begin
      Done := True;
      SetLength(Result, i - 1);
    end;
    Dec(i);
  end;
end;

function WideStripRemoteComputer(const UNCPath: string): string;
  // Strips the \\RemoteComputer\ part of an UNC path
var
  Head: PWideChar;
begin
  Result := '';
  if IsUNCPath(UNCPath) then
  begin
    Result := '';
    if IsUNCPath(UNCPath) then
    begin
      Result := UNCPath;
      Head := @Result[1];
      Head := Head + 2;    // Skip past the '\\'
      Head := SysUtils.StrScan(Head, WideChar('\'));
      if Assigned(Head) then
      begin
        Head := Head + 1;
        Move(Head[0], Result[1], (lstrlenW(Head) + 1) * 2);
      end;
      SetLength(Result, lstrlenW(PWideChar(Result)));
    end;
  end;
end;

function WideStripTrailingBackslash(const S: string; Force: Boolean = False): string;
begin
  Result := S;
  if Result <> '' then
  begin
    // Works with FilePaths and FTP Paths
    if (Result[ Length(Result)] = WideChar('/')) or (Result[ Length(Result)] = WideChar('\')) then
      if not WideIsDrive(Result) or Force then // Don't strip off if is a root drive
        SetLength(Result, Length(Result) - 1);
  end;
end;

function WideStripLeadingBackslash(const S: string): string;
begin
  Result := S;
  if Result <> '' then
  begin
    if (S[1] = '\') and (Length(S) > 1) then
      Result := PWideChar( @S[2])
    else
      Result := ''
  end;
end;

function WideShellExecute(hWnd: HWND; Operation, FileName, Parameters, Directory: string; ShowCmd: Integer = SW_NORMAL): HINST;
var
  PW, DW: PWideChar;
begin
  PW := nil;
  DW := nil;
  if Parameters <> '' then
    PW := PWideChar(Parameters);
  if Directory <> '' then
    DW := PWideChar(Directory);
  Result := ShellExecute(hWnd, PWideChar(Operation), PWideChar(FileName), PW, DW, SW_NORMAL)
end;

procedure WideShowMessage(Window: HWND; ACaption, AMessage: string);
begin
  Windows.MessageBox(Window, PWideChar(AMessage), PWideChar(ACaption), MB_ICONEXCLAMATION or MB_OK)
end;

function DiskInDrive(C: AnsiChar): Boolean;
var
  OldErrorMode: Integer;
begin
  C := UpCase(C);
  if C in ['A'..'Z'] then
  begin
    OldErrorMode := SetErrorMode(SEM_FAILCRITICALERRORS);
    Result :=  DiskFree(Ord(C) - Ord('A') + 1) > -1;
    SetErrorMode(OldErrorMode);
  end else
    Result := False
end;

function WideStrIComp(Str1, Str2: PWideChar): Integer;
// Insensitive case comparison
begin
  Result := lstrcmpi(Str1, Str2);
end;

function WideStrComp(Str1, Str2: PWideChar): Integer;
// Sensitive case comparison
begin
  Result := lstrcmp(Str1, Str2);
end;

function ShortenStringEx(DC: HDC; const S: string; Width: Integer; RTL: Boolean;
  EllipsisPlacement: TShortenStringEllipsis): string;
// Shortens the passed string and inserts ellipsis "..." in the requested place.
// This is not a fast function but it is clear how it works.  Also the RTL implmentation
// is still being understood.
var
  Len: Integer;
  EllipsisWidth: Integer;
  TargetString: WideString;
  Tail, Head, Middle: PWideChar;
  L, ResultW: integer;
begin
  Len := Length(S);
  if (Len = 0) or (Width <= 0) then
    Result := ''
  else begin
    // Determine width of triple point using the current DC settings
    TargetString := '...';
    EllipsisWidth := TextExtentW(PWideChar(TargetString), DC).cx;

    if Width <= EllipsisWidth then
      Result := ''
    else begin
      TargetString := S;
      Head := PWideChar(TargetString);
      Tail := Head;
      Inc(Tail, lstrlenW(PWideChar(TargetString)));
      case EllipsisPlacement of
        sseEnd:
          begin
            L := EllipsisWidth + TextExtentW(PWideChar(TargetString), DC).cx;
            while (L > Width) do
            begin
              Dec(Tail);
              Tail^ := WideNull;
              L := EllipsisWidth + TextExtentW(PWideChar(TargetString), DC).cx;
            end;
            Result := PWideChar(TargetString) + '...';
          end;
        sseFront:
          begin
            L := EllipsisWidth + TextExtentW(PWideChar(TargetString), DC).cx;
            while (L > Width) do
            begin
              Inc(Head);
              L := EllipsisWidth + TextExtentW(PWideChar(Head), DC).cx;
            end;
            Result := '...' + PWideChar(Head);
          end;
        sseMiddle:
          begin
            L := EllipsisWidth + TextExtentW(PWideChar(TargetString), DC).cx;
            while (L > Width div 2 - EllipsisWidth div 2) do
            begin
              Dec(Tail);
              Tail^ := WideNull;
              L := TextExtentW(PWideChar(TargetString), DC).cx;
            end;
            Result := PWideChar(TargetString) + '...';
            ResultW := TextExtentW(PWideChar(Result), DC).cx;

            TargetString := S;
            Head := PWideChar(TargetString);
            Middle := Head;
            Inc(Middle, lstrlenW(PWideChar(Result)) - 3); // - 3 for ellipsis letters
            Tail := Head;
            Inc(Tail, lstrlenW(PWideChar(TargetString)));
            Head := Tail - 1;

            L := ResultW + TextExtentW(Head, DC).cx;
            while (L < Width) and (Head >= Middle) do
            begin
              Dec(Head);
              L := ResultW + TextExtentW(PWideChar(Head), DC).cx;
            end;
            Inc(Head);

            Result := Result + Head;
          end;
        sseFilePathMiddle:
          begin
            Head := Tail - 1;
            L := EllipsisWidth + TextExtentW(Head, DC).cx;
            while (Head^ <> '\') and (Head <> @TargetString[1]) and (L < Width) do
            begin
              Dec(Head);
              L := EllipsisWidth + TextExtentW(Head, DC).cx;
            end;
            if Head^ <> '\' then
              Inc(Head);
            Result := '...' + Head;

            ResultW := TextExtentW(PWideChar(Result), DC).cx;

            Head^ := WideNull;
            SetLength(TargetString, lstrlenW(PWideChar(TargetString)));

            Head := PWideChar(TargetString);
            Tail := Head;
            Inc(Tail, lstrlenW(Head));

            L := ResultW + TextExtentW(PWideChar(TargetString), DC).cx;
            while (L > Width) and (Tail > @TargetString[1]) do
            begin
              Dec(Tail);
              Tail^ := WideNull;
              L := ResultW + TextExtentW(PWideChar(TargetString), DC).cx;
            end;

            Result := Head + Result;
          end;
      end;

      // Windows 2000 automatically switches the order in the string. For every other system we have to take care.
 {     if IsWin2000 or not RTL then
        Result := Copy(S, 1, N - 1) + '...'
      else
        Result := '...' + Copy(S, 1, N - 1);       }
    end;
  end;
end;

function WindowsDirectory: string;
var
  Len: integer;
begin
  Result := '';
  Len := GetWindowsDirectory(PWideChar(Result), 0);
  if Len > 0 then
  begin
    SetLength(Result, Len - 1);
    GetWindowsDirectory(PWideChar(Result), Len);
  end
end;

function ModuleFileName(PathOnly: Boolean = True): string;
var
  BufferW: array[0..MAX_PATH] of WideChar;
begin
  FillChar(BufferW, SizeOf(BufferW), #0);
  if GetModuleFileName(0, BufferW, SizeOf(BufferW)) > 0 then
  begin
    if PathOnly then
      Result := ExtractFileDir(BufferW)
    else
      Result := BufferW;
  end
end;

function PIDLToPath(PIDL: PItemIDList): string;
var
  PIDLMgr: TCommonPIDLManager;
  LastID: PItemIDList;
  LastCB: Word;
  Desktop, Folder: IShellFolder;
  StrRet: TStrRet;
begin
  Result := '';
  if Assigned(PIDL) then
  begin
    FillChar(StrRet, SizeOf(StrRet), #0);
    PIDLMgr := TCommonPIDLManager.Create;
    try
      SHGetDesktopFolder(Desktop);
      if PIDLMgr.IsDesktopFolder(PIDL) then
      begin
        if Succeeded(Desktop.GetDisplayNameOf(PIDL, SHGDN_FORPARSING, StrRet)) then
          Result := StrRetToStr(StrRet, PIDL);
      end else
      begin
        PIDLMgr.StripLastID(PIDL, LastCB, LastID);
        if Assigned(LastID) then
          try
            if Succeeded(Desktop.BindToObject(PIDL, nil, IShellFolder, pointer(Folder))) then
            begin
              LastID.mkid.cb := LastCB;
              if Succeeded(Folder.GetDisplayNameOf(LastID, SHGDN_FORPARSING, StrRet)) then
                 Result := StrRetToStr(StrRet, LastID);
            end
          finally
            LastID.mkid.cb := LastCB;
          end
      end
    finally
      PIDLMgr.Free
    end
  end
end;

function ShortFileName(const FileName: string): string;
var
  BufferW: array[0..MAX_PATH] of WideChar;
begin
  if GetShortPathName(PWideChar(FileName), BufferW, SizeOf(BufferW)) > 0 then
    Result := BufferW
  else
    Result := FileName;
end;

function ShortPath(const Path: string): string;
begin
  Result := ShortFileName(Path)
end;

procedure LoadWideString(S: TStream; var Str: string);
var
  Count: Integer;
begin
  S.Read(Count, SizeOf(Integer));
  SetLength(Str, Count);
  S.Read(PWideChar(Str)^, Count * 2)
end;

function UsesAlphaChannel(Image32: TBitmap): Boolean;
var
  I, N: Integer;
  LineDeltaImage32, PixelDeltaImage32: Integer;
  Alpha: Byte;
  PLineImage32, PPixelImage32: PByte;
begin
  Result := False;
  if Assigned(Image32) and (Image32.PixelFormat = pf32Bit) and (Image32.Height > 1) then
  begin
      LineDeltaImage32 := Integer( DWORD( Image32.ScanLine[1]) - DWORD( Image32.ScanLine[0]));
      PixelDeltaImage32 := SizeOf(TRGBQuad);
      PLineImage32 := Image32.ScanLine[0];

      for I := 0 to Image32.Height - 1 do
      begin
        PPixelImage32 := PLineImage32;

        for N := 0 to Image32.Width - 1 do
        begin
          // Source GetColorValues ; Profiled = ~24-30% of time
          Alpha := (PDWORD( PPixelImage32)^ and $FF000000) shr 24;
          Result := Alpha <> 0;
          if Result then
            Exit;
          Inc(PPixelImage32, PixelDeltaImage32);
        end;
        Inc(PLineImage32, LineDeltaImage32);
      end;
  end
end;

procedure SaveWideString(S: TStream; Str: string);
var
  Count: Integer;
begin
  Count := Length(Str);
  S.Write(Count, SizeOf(Integer));
  S.Write(PWideChar(Str)^, Count * 2)
end;

function ProperRect(Rect: TRect): TRect;
// Makes sure a rectangle's left is less than its right and its top is less than its bottom
var
  Temp: integer;
begin
  Result := Rect;
  if Result.Right < Result.Left then
  begin
    Temp := Result.Right;
    Result.Right := Rect.Left;
    Result.Left := Temp;
  end;
  if Rect.Bottom < Rect.Top then
  begin
    Temp := Result.Top;
    Result.Top := Rect.Bottom;
    Result.Bottom := Temp;
  end
end;

function DragDetectPlus(Handle: HWND; Pt: TPoint): Boolean;
//  Replacement for DragDetect API which is buggy.
//    Pt is in Client Coords of the Handle window
var
  DragRect: TRect;
  Msg: TMsg;
  TestPt: TPoint;
  HadCapture, Done: Boolean;
begin
  Result := False;
  Done := False;
  HadCapture := GetCapture = Handle;
  if (not ClientToScreen(Handle, Pt)) then
    Exit;
  DragRect.TopLeft := Pt;
  DragRect.BottomRight := Pt;
  InflateRect(DragRect, GetSystemMetrics(SM_CXDRAG), GetSystemMetrics(SM_CYDRAG));
  SetCapture(Handle);
  try
    while (not Result) and (not Done) do
      if (PeekMessage(Msg, Handle, 0,0, PM_REMOVE)) then
      begin
        case (Msg.message) of
          WM_MOUSEMOVE:
          begin
            TestPt := Msg.Pt;
            // Not sure why this works.  The Message point "should" be in client
            // coordinates but seem to be screen
     //       Windows.ClientToScreen(Msg.hWnd, TestPt);
            Result := not(PtInRect(DragRect, TestPt));
          end;
          WM_RBUTTONUP,
          WM_LBUTTONUP,
          WM_CANCELMODE,
          WM_LBUTTONDBLCLK,
          WM_MBUTTONUP:
            begin
              // Let the window get these messages after we have ended our
              // local message loop
              PostMessage(Msg.hWnd, Msg.message, Msg.wParam, Msg.lParam);
              Done := True;
            end;
          WM_QUIT:
            begin
              PostQuitMessage(Msg.wParam);
              Done := True;
            end
        else
          TranslateMessage(Msg);
          DispatchMessage(Msg)
        end
      end else
        Sleep(0);
  finally
    ReleaseCapture;
    if HadCapture then
      Mouse.Capture := Handle;
  end;
end;

procedure FreeMemAndNil(var P: Pointer);
{ Frees the memeory allocated with GetMem and nils the pointer                  }
var
  Temp: Pointer;
begin
  Temp := P;
  P := nil;
  FreeMem(Temp);
end;

function IsRectNull(ARect: TRect): Boolean;
begin
  Result := EqualRect(ARect, Rect(0, 0, 0, 0))
end;

function IsUnicode: Boolean;
begin
  Result := Win32Platform = VER_PLATFORM_WIN32_NT
end;

function IsWinNT: Boolean;
begin
  Result := Win32Platform = VER_PLATFORM_WIN32_NT
end;

function IsWin95_SR1: Boolean;
begin
  Result := False;
  if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
    Result :=  ((Win32MajorVersion = 4) and
                (Win32MinorVersion = 0) and
                (LoWord(Win32BuildNumber) <= 1080))
end;

function IsWinME: Boolean;
begin
  Result := False;
  if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
    Result := Win32BuildNumber >= $045A0BB8
end;

function IsWinNT4: Boolean;
begin
  Result := ( Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion < 5)
end;

function IsWin2000: Boolean;
begin
  Result := ( Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 5) and (Win32MinorVersion >= 0)
end;

function IsWinXP: Boolean;
begin
  Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 5) and (Win32MinorVersion > 0)
end;

function IsWinVista: Boolean;
begin
  Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 6) and (Win32MinorVersion >= 0)
end;

function IsWin7: Boolean;
begin
  Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 6) and (Win32MinorVersion > 0)
end;

function RectHeight(R: TRect): Integer;
begin
  Result := R.Bottom - R.Top
end;

function RectToStr(R: TRect): string;
begin
  Result := 'Left = ' + IntToStr(R.Left) +
            ' Top = ' + IntToStr(R.Top) +
            ' Right = ' + IntToStr(R.Right) +
            ' Bottom = ' + IntToStr(R.Bottom)
end;

function RectToSquare(R: TRect): TRect;
// Takes the passed rectangle and makes it square based on the longest dimension
begin
  if RectWidth(R) > RectHeight(R) then
    R.Right := R.Left + RectHeight(R)
  else
  if RectHeight(R) > RectWidth(R) then
  begin
    R.Bottom := R.Top + RectWidth(R);
  end else
    Result := R
end;

function RectWidth(R: TRect): Integer;
begin
  Result := R.Right - R.Left
end;

function ContainsRect(OuterRect, InnerRect: TRect): Boolean;
//
//  Returns true if the InnerRect is completely contained within the
//  OuterRect
//
begin
  Result := (InnerRect.Left >= OuterRect.Left) and
            (InnerRect.Right <= OuterRect.Right) and
            (InnerRect.Top >= OuterRect.Top) and
            (InnerRect.Bottom <= OuterRect.Bottom)
end;

function KeyDataToShiftState(KeyData: Longint): TShiftState;
begin
  Result := [];
  if GetKeyState(VK_SHIFT) < 0 then
    Include(Result, ssShift);
  if GetKeyState(VK_CONTROL) < 0 then
    Include(Result, ssCtrl);
  if KeyData and $20000000 <> 0 then
    Include(Result, ssAlt);
end;

function KeyDataAsyncToShiftState: TShiftState;
begin
  Result := [];
  if AsyncShiftDown then
    Include(Result, ssShift);
  if AsyncControlDown then
    Include(Result, ssCtrl);
  if AsyncAltDown then
    Include(Result, ssAlt);
end;

function DropEffectToDropEffectState(Effect: Integer): TCommonDropEffect;
begin
  Result := cdeNone;
  if Effect and DROPEFFECT_COPY <> 0 then
    Result := cdeCopy
  else
  if Effect and DROPEFFECT_MOVE <> 0 then
    Result := cdeMove
  else
  if Effect and DROPEFFECT_LINK <> 0 then
    Result := cdeLink
end;

function DropEffectStateToDropEffect(Effect: TCommonDropEffect): Integer;
begin
  Result := 0;
  if cdeCopy = Effect then
    Result := Result or DROPEFFECT_COPY
  else
  if cdeMove = Effect then
    Result := Result or DROPEFFECT_MOVE
  else
  if cdeLink = Effect then
    Result := Result or DROPEFFECT_LINK
end;

function DropEffectToDropEffectStates(Effect: Integer): TCommonDropEffects;
begin
  Result := [];
  if (Effect = 0) or (Longword(Effect) = DROPEFFECT_SCROLL) then
    Include(Result, cdeNone);
  if Effect and DROPEFFECT_COPY <> 0 then
    Include(Result, cdeCopy);
  if Effect and DROPEFFECT_MOVE <> 0 then
    Include(Result, cdeMove);
  if Effect and DROPEFFECT_LINK <> 0 then
    Include(Result, cdeLink);
  if Effect and DROPEFFECT_SCROLL <> 0 then
    Include(Result, cdeScroll)
end;

function DropEffectStatesToDropEffect(Effect: TCommonDropEffects): Integer;
begin
  Result := 0;
  if cdeCopy in Effect then
    Result := Result or DROPEFFECT_COPY;
  if cdeMove in Effect then
    Result := Result or DROPEFFECT_MOVE;
  if cdeLink in Effect then
    Result := Result or DROPEFFECT_LINK;
  if cdeScroll in Effect then
    Result := Result or Integer(DROPEFFECT_SCROLL);
end;

function KeyToKeyStates(Keys: Word): TCommonKeyStates;
begin
  Result := [];
  if Keys and MK_CONTROL <> 0 then
    Include(Result, cksControl);
  if Keys and MK_LBUTTON <> 0 then
    Include(Result, cksLButton);
  if Keys and MK_MBUTTON <> 0 then
    Include(Result, cksMButton);
  if Keys and MK_RBUTTON <> 0 then
    Include(Result, cksRButton);
  if Keys and MK_SHIFT <> 0 then
    Include(Result, cksShift);
  if Keys and MK_ALT <> 0 then
    Include(Result, cksAlt);
  if Keys and MK_BUTTON <> 0 then
    Include(Result, cksButton);
end;

function KeyStatesToMouseButton(Keys: Word): TCommonMouseButton;
begin
  if Keys and MK_LBUTTON <> 0 then
    Result := cmbLeft
  else
  if Keys and MK_MBUTTON <> 0 then
    Result := cmbMiddle
  else
  if Keys and MK_RBUTTON <> 0 then
    Result := cmbRight
  else
    Result := cmbNone
end;

function KeyStatesToKey(Keys: TCommonKeyStates): Longword;
begin
  Result := 0;
  if cksControl in Keys then
    Result := Result or MK_CONTROL;
  if cksLButton in Keys then
    Result := Result or MK_LBUTTON;
  if cksMButton in Keys then
    Result := Result or MK_MBUTTON;
  if cksRButton in Keys then
    Result := Result or MK_RBUTTON;
  if cksShift in Keys then
    Result := Result or MK_SHIFT;
  if cksAlt in Keys then
    Result := Result or MK_ALT;
end;

function KeyStateToDropEffect(Keys: TCommonKeyStates): TCommonDropEffect;
begin
  Result := cdeMove;  // The default
  if (cksControl in Keys) and not ((cksShift in Keys) or (cksAlt in Keys)) then
    Result := cdeCopy
  else
  if ((cksAlt in Keys) and not ((cksShift in Keys) or (cksControl in Keys))) or
     ((cksShift in Keys) and (cksControl in Keys)) then
    Result := cdeLink
end;

function KeyStateToMouseButton(KeyState: TCommonKeyStates): TCommonMouseButton;
begin
  if KeyState * [cksLButton] <> [] then
    Result := cmbLeft
  else
  if KeyState * [cksRButton] <> [] then
    Result := cmbRight
  else
  if KeyState * [cksMButton] <> [] then
    Result := cmbMiddle
  else
    Result := cmbNone
end;

function FileIconInit(FullInit: BOOL): BOOL; stdcall;
// Forces the system to load all Images into the ImageList. Normally with NT
// system icons are only loaded as needed.
type
  TFileIconInit = function(FullInit: BOOL): BOOL; stdcall;
var
  ShellDLL: HMODULE;
  PFileIconInit: TFileIconInit;
begin
  Result := False;
  ShellDLL := CommonLoadLibrary(Shell32);
  PFileIconInit := GetProcAddress(ShellDLL, PAnsiChar(660));
  if (Assigned(PFileIconInit)) then
    Result := PFileIconInit(FullInit);
end;

function SHGetImageList(iImageList: Integer; const RefID: TGUID; out ppvOut): HRESULT; stdcall;
// Retrieves the system ImageList interface
var
  ShellDLL: HMODULE;
  ImageList: TSHGetImageList;
begin
  Result := E_NOTIMPL;
 //   if GetFileVersion(comctl32) >= $00060000 then
  begin
    ShellDLL := CommonLoadLibrary(Shell32);
    ImageList := GetProcAddress(ShellDLL, PAnsiChar(727));
    if (Assigned(ImageList)) then
      Result := ImageList(iImageList, RefID, ppvOut);
  end
end;

function Size(cx, cy: Integer): TSize;
begin
  Result.cy := cy;
  Result.cx := cx
end;

function ShortenTextW(DC: hDC; TextToShorten: string; MaxSize: Integer): string;
// Shortens the passed string in such a way that if it does not fit in the MaxSize
// (in Pixels) a "..." is inserted at the correct place where the new string fixs
// in MaxSize
var
  Size: TSize;
  EllipsisSize: TSize;
  StrLen, Middle, Low, High{, LastHigh}: Cardinal;
begin
  if TextToShorten <> '' then
  begin
    GetTextExtentPoint32W(DC, PWideChar(TextToShorten), Length(TextToShorten), Size);
    GetTextExtentPoint32W(DC, '...', 3, EllipsisSize);
    StrLen := Length(TextToShorten);
    if Size.cx > MaxSize then
    begin
      Low:=0;
      High:=StrLen-1;
      while (Low<High) do
        begin
          Middle:=(Low+High+1) shr 1;
          GetTextExtentPoint32W(DC, @TextToShorten[1], Middle, Size);
          Size.cx := Size.cx + EllipsisSize.cx;
          if (Size.cx<=MaxSize) then
            Low:=Middle
          else
            High:=Middle-1;
        end;

      SetLength(TextToShorten, Low);
      if Low > 0 then
        Result := TextToShorten + '...'
      else
        Result := '...'
    end else
      Result := TextToShorten
  end else
    Result := '';
end;

// Solerman's version
function SplitTextW(DC: hDC; TextToSplit: string; MaxWidth: Integer;
   var Buffer: TCommonWideCharArray; MaxSplits: Integer): Integer;
// Takes the passed string and breaks it up so each piece fits within the MaxWidth
// The function detects any LF/CR pairs and treats them as one break if CR or LF
// is defined as a break character.
// The Buffer is a set of NULL terminated strings for each line, with the last
// one being terminated with a double NULL.  Much like the SHFileOperation API
// This makes it ready to use to pass the strings to ExtTextOutW in a loop
// If the buffer is too small the Result will be false
// If MaxSplits = -1 then the function splits as many time as necessary
// The Return is the total number of lines the passed text was split into
var
   Head, Tail, LastBreakChar, BufferHead: PWideChar;
//   TailA: String;
   Size: TSize;
   LineWidth, SplitCount, Len: Integer;
   TextMetrics: TTextMetric;
begin
   Result := 0;
   if MaxWidth = 0 then
   begin
     SetLength(Buffer, 2);
     Buffer[0] := WideNull;
     Buffer[1] := WideNull;
   end else
   begin
     GetTextMetrics(DC, TextMetrics);

     // Can get into deep trouble if a single letter won't fit in the MaxWidth
     if TextMetrics.tmMaxCharWidth > MaxWidth then
     begin
       Len := Length(TextToSplit);
       SetLength(Buffer, Len + 2);
       if Len > 0 then
       begin
         Head := @TextToSplit[1];
         CopyMemory(Buffer, Head, Len*2);
         Result := 1;
       end;
       Buffer[Len] := #0;
       Buffer[Len + 1] := #0;
     end else
     begin
       FillChar(Size, SizeOf(Size), #0);
       // Arbitrary size that should be ok in most instances.  This will be enough space
       // for 127 lines
       SetLength(Buffer, Length(TextToSplit) + 128);
       BufferHead := @Buffer[0];
       Head := PWideChar(TextToSplit);
       Tail := Head;
       SplitCount := 0;
       while (Tail^ <> #0) and ((SplitCount < MaxSplits) or (MaxSplits = -1)) do
       begin
         LineWidth := 0;
         LastBreakChar := nil;
         while (Tail^ <> WideNull) and (LineWidth <= MaxWidth) and (Tail^ <> WideCR) and (Tail^ <> WideLF) do
         begin
           GetTextExtentPoint32W(DC, PWideChar(Tail), 1, Size);
           Inc(LineWidth, Size.cx);

           Inc(Tail);

           if (LineWidth <= MaxWidth) and (Tail^ = WideSpace) then
             LastBreakChar := Tail;
         end;

         if (LineWidth > MaxWidth) then
         begin
           // We overran the MaxWidth if entering this block

           // Over ran the line unless it exactly fits
           if Assigned(LastBreakChar) and (LineWidth > MaxWidth) then
             // We have word break to go back to
             Tail := LastBreakChar
           else
             // Special case, the Tail is the end of the Text to split
             if {(Tail^ <> WideNull) and} ((SplitCount + 1 < MaxSplits) or (MaxSplits < 0)) then
               Dec(Tail);
         end;
         Inc(SplitCount);
         // If we reach the MaxSplits make the last line be the rest of the text.
         if (SplitCount > 0) and (SplitCount = MaxSplits) then
           Inc(Tail, lstrlenW(Tail));
         CopyMemory(BufferHead, Head, (Tail-Head)*2);
         Inc(BufferHead, Tail-Head);
         BufferHead^ := WideNull;
         Inc(BufferHead);
         Inc(Result);

         while (Tail^ <> #0) and ((Tail^ = WideCR) or (Tail^ = WideLF) or (Tail^ = WideSpace)) do
           Inc(Tail);
         Head := Tail
       end
     end
   end
end;

function StrCopyW(Dest, Source: PWideChar): PWideChar;
// copies Source to Dest and returns Dest
{$ifdef CPUX64}
begin
  Result :=  StrCopy(Dest, Source);
end;
{$else}
asm
       PUSH    EDI
       PUSH    ESI
       MOV     ESI, EAX
       MOV     EDI, EDX
       MOV     ECX, 0FFFFFFFFH
       XOR     AX, AX
       REPNE   SCASW
       NOT     ECX
       MOV     EDI, ESI
       MOV     ESI, EDX
       MOV     EDX, ECX
       MOV     EAX, EDI
       SHR     ECX, 1
       REP     MOVSD
       MOV     ECX, EDX
       AND     ECX, 1
       REP     MOVSW
       POP     ESI
       POP     EDI
end;
{$endif CPUX64}

//----------------------------------------------------------------------------------------------------------------------

function BrightenColor(const RGB: TCommonRGB; Amount: Double): TCommonRGB;
var
  HLS: TCommonHLS;

begin
  HLS := RGBToHLS(RGB);
  HLS.L := (1 + Amount) * HLS.L;
  Result := HLSToRGB(HLS);
end;

function DarkenColor(const RGB: TCommonRGB; Amount: Double): TCommonRGB;
var
  HLS: TCommonHLS;

begin
  HLS := RGBToHLS(RGB);
  // Darken means to decrease luminance.
  HLS.L := (1 - Amount) * HLS.L;
  Result := HLSToRGB(HLS);
end;

function RGBToHLS(const RGB: TCommonRGB): TCommonHLS;

// Converts from RGB to HLS.
// Input parameters and result values are all in the range 0..1.
// Note: Hue is normalized so 360� corresponds to 1.

var
  Delta,
  Max,
  Min:  Double;

begin
  with RGB, Result do
  begin
    Max := MaxValue([R, G, B]);
    Min := MinValue([R, G, B]);

    L := (Max + Min) / 2;

    if Max = Min then
    begin
      // Achromatic case.
      S := 0;
      H := 0; // Undefined
    end
    else
    begin
      Delta := Max - Min;

      if L < 0.5 then
        S := Delta / (Max + Min)
      else
        S := Delta / (2 - (Max + Min));

      if R = Max then
        H := (G - B) / Delta
      else
        if G = Max then
          H := 2 + (B - R) / Delta
        else
          if B = Max then
            H := 4 + (R - G) / Delta;

      H := H / 6;
      if H < 0 then
        H := H + 1;
    end
  end;
end;

function HLSToRGB(const HLS: TCommonHLS): TCommonRGB;

// Converts from HLS (hue, luminance, saturation) to RGB.
// Input parameters and result values are all in the range 0..1.
// Note: Hue is normalized so 360� corresponds to 1.

  //--------------- local function --------------------------------------------

  function HueToRGB(m1, m2, Hue: Double): Double;

  begin
    if Hue > 1 then
      Hue := Hue - 1
    else
      if Hue < 0 then
        Hue := Hue + 1;

    if 6 * Hue < 1 then
      Result := m1 + (m2 - m1) * Hue * 6
    else
      if 2 * Hue < 1 then
        Result := m2
      else
        if 3 * Hue < 2 then
          Result := m1 + (m2 - m1) * (2 / 3 - Hue) * 6
        else
          Result := m1;
  end;

  //--------------- end local function ----------------------------------------

var
  m1, m2: Double;

begin
  with HLS, Result do
  begin
    if S = 0 then
    begin
      // Achromatic case (no hue).
      R := L;
      G := L;
      B := L;
    end
    else
    begin
      if L <= 0.5 then
        m2 := L * (S + 1)
      else
        m2 := L + S - L * S;
      m1 := 2 * L - m2;

      R := HueToRGB(m1, m2, H + 1 / 3);
      G := HueToRGB(m1, m2, H);
      B := HueToRGB(m1, m2, H - 1 / 3)
    end;
  end;
end;

function MakeTRBG(Color: TColor): TCommonRGB;
var
  RGB: Longint;
begin
  RGB := ColorToRGB(Color);
  Result.B := GetBValue(RGB) / 255;
  Result.G := GetGValue(RGB) / 255;
  Result.R := GetRValue(RGB) / 255;
end;

function MakeTColor(RGB: TCommonRGB): TColor;
begin
  Result := TColor(MakeColorRef(RGB));
end;

function MakeColorRef(RGB: TCommonRGB; Gamma: Double = 1): COLORREF;
// Converts a floating point RGB color to an 8 bit color reference as used by Windows.
// The function takes care not to produce out-of-gamut colors and allows to apply an optional gamma correction
// (inverse gamma).
begin
  GammaCorrection(RGB, Gamma);
  with RGB do
    Result := Windows.RGB(Round(R * 255), Round(G * 255), Round(B * 255));
end;

procedure GammaCorrection(var RGB: TCommonRGB; Gamma: Double);

// Computes the gamma corrected RGB color and ensures the result is in-gamut.

begin
  if Gamma <> 1 then
  begin
    Gamma := 1 / Gamma;
    with RGB do
    begin
      if R > 0 then
        R := Power(R, Gamma)
      else
        R := 0;
      if G > 0 then
        G := Power(G, Gamma)
      else
        G := 0;
      if B > 0 then
        B := Power(B, Gamma)
      else
        B := 0;
    end;
  end;

  MakeSafeColor(RGB);
end;

function MakeSafeColor(var RGB: TCommonRGB): Boolean;
// Ensures the given RGB color is in-gamut, that is, no component is < 0 or > 1.
// Returns True if the color had to be adjusted to be in-gamut, otherwise False.
begin
  Result := False;

  if RGB.R < 0 then
  begin
    Result := True;
    RGB.R := 0;
  end;
  if RGB.R > 1 then
  begin
    Result := True;
    RGB.R := 1;
  end;

  if RGB.G < 0 then
  begin
    Result := True;
    RGB.G := 0;
  end;
  if RGB.G > 1 then
  begin
    Result := True;
    RGB.G := 1;
  end;

  if RGB.B < 0 then
  begin
    Result := True;
    RGB.B := 0;
  end;
  if RGB.B > 1 then
  begin
    Result := True;
    RGB.B := 1;
  end;
end;

  function UpsideDownDIB(Bits: TBitmap): Boolean;
var
  OldRGB, TempRGB: LongInt;
  P: PLongInt;
begin
  Result := False;
  Assert(Bits.PixelFormat = pf32Bit, 'UpsideDownDIB only works with 32 bit bitmaps');
  if Bits.PixelFormat = pf32Bit then
  begin
    P := Bits.ScanLine[0];
    OldRGB := P^;
    Result := True;
    // if Equal then we can can't be sure if upsidedown
    if (P^ and $00FFFFFF) = (ColorToRGB(Bits.Canvas.Pixels[0, 0]) and $00FFFFFF) then
    begin
      // Flip the pixel bits
      TempRGB := not OldRGB and $00FFFFFF;
      Bits.Canvas.Pixels[0, 0] := TempRGB;

      Result := (P^ and $00FFFFFF) <> (ColorToRGB(Bits.Canvas.Pixels[0, 0]) and $00FFFFFF);
      P^ := OldRGB
    end
  end
end;

procedure ActivateTopLevelWindow(Child: HWND);
var
  Parent: HWND;
  Style: LongWord;
begin
  Parent := GetParent(Child);
  while Parent <> 0 do
  begin
    Style := GetWindowLong(Parent, GWL_STYLE);
    if ((Style and WS_POPUP <> 0)) or (Style and WS_CHILD = 0) then
    begin
      BringWindowToTop(Parent);
      Exit
    end;
    Parent := GetParent(Parent);
  end
end;

function IsVScrollVisible(Wnd: HWnd): Boolean;
begin
  Result := WS_VSCROLL and GetWindowLong(Wnd, GWL_STYLE) <> 0
end;

function IsHScrollVisible(Wnd: HWnd): Boolean;
begin
  Result := WS_HSCROLL and GetWindowLong(Wnd, GWL_STYLE) <> 0
end;

function HScrollbarHeightRuntime(Wnd: HWnd): Integer;
begin
  Result := 0;
  if IsHScrollVisible(Wnd) then
    Result := HScrollbarHeight;
end;

function VScrollbarWidthRuntime(Wnd: HWnd): Integer;
begin
  Result := 0;
  if IsVScrollVisible(Wnd) then
    Result := VScrollbarWidth;
end;

function HScrollbarHeight: Integer;
begin
  Result := GetSystemMetrics(SM_CYHSCROLL);
end;

function VScrollbarWidth: Integer;
begin
  Result := GetSystemMetrics(SM_CXVSCROLL);
end;

procedure LoadWideFunctions;
begin
  if not WideFunctionsLoaded then
  begin
    // We can be sure these are already loaded.  This keeps us from having to
    // reference count when VSTools is being used in an OCX
    Shell32Handle := GetModuleHandle(Shell32);
    Kernel32Handle := GetModuleHandle(Kernel32);
    AdvAPI32Handle := GetModuleHandle(AdvAPI32);
    ShlwapiHandle := CommonLoadLibrary(Shlwapi);
    UserEnvHandle := CommonLoadLibrary('Userenv.dll');

    if ShlwapiHandle <> 0 then
    begin
      PathMatchSpecW_MP := GetProcAddress(ShlwapiHandle, 'PathMatchSpecW');
    end;

    ExpandEnvironmentStringsForUserW_MP := GetProcAddress(UserEnvHandle, 'ExpandEnvironmentStringsForUserW');

    // NTFS Volume (Junction) only functions
    Wow64RevertWow64FsRedirection_MP := GetProcAddress(Kernel32Handle, 'Wow64RevertWow64FsRedirection');
    Wow64DisableWow64FsRedirection_MP := GetProcAddress(Kernel32Handle, 'Wow64DisableWow64FsRedirection');
    Wow64EnableWow64FsRedirection_MP := GetProcAddress(Kernel32Handle, 'Wow64EnableWow64FsRedirection');   // Don't use this per M$, use the above two to disable and rever

    NetAPI32Handle := CommonLoadLibrary('NETAPI32.DLL');
    if NetAPI32Handle <> 0 then
    begin
      NetShareEnumW_MP := GetProcAddress(NetAPI32Handle, 'NetShareEnum');
      NetApiBufferFree_MP := GetProcAddress(NetAPI32Handle, 'NetApiBufferFree');
    end;

    // Windows 7 stuff
    SetCurrentProcessExplicitAppUserModelID_MP := GetProcAddress(Shell32Handle, 'SetCurrentProcessExplicitAppUserModelID');
    SHGetPropertyStoreForWindow_MP := GetProcAddress(Shell32Handle, 'SHGetPropertyStoreForWindow');

    CDefFolderMenu_Create2_MP := GetProcAddress(Shell32Handle, PAnsiChar(701));

    WideFunctionsLoaded := True
  end
end;


{ TCallBackStub }

  // Helpers to create a callback function out of a object method
{ ----------------------------------------------------------------------------- }
{ This is a piece of magic by Jeroen Mineur.  Allows a class method to be used  }
{ as a callback. Create a stub using CreateStub with the instance of the object }
{ the callback should call as the first parameter and the method as the second  }
{ parameter, ie @TForm1.MyCallback or declare a type of object for the callback }
{ method and then use a variable of that type and set the variable to the       }
{ method and pass it:                                                           }
{ 64-bit code by Andrey Gruzdev                                                 }
{                                                                               }
{ type                                                                          }
{   TEnumWindowsFunc = function (AHandle: hWnd; Param: lParam): BOOL of object; stdcall; }
{                                                                               }
{  TForm1 = class(TForm)                                                        }
{  private                                                                      }
{    function EnumWindowsProc(AHandle: hWnd; Param: lParam): BOOL; stdcall;     }
{  end;                                                                         }
{                                                                               }
{  var                                                                          }
{    MyFunc: TEnumWindowsFunc;                                                  }
{    Stub: ICallbackStub;                                                       }
{  begin                                                                        }
{    MyFunct := EnumWindowsProc;                                                }
{    Stub := TCallBackStub.Create(Self, MyFunct, 2);                            }
{     ....                                                                      }
{     ....                                                                      }
{  Now Stub.StubPointer can be passed as the callback pointer to any windows API}
{  The Stub will be automatically disposed when the interface variable goes out }
{  of scope                                                                     }
{ ----------------------------------------------------------------------------- }
{$IFNDEF CPUX64}
const
  AsmPopEDX = $5A;
  AsmMovEAX = $B8;
  AsmPushEAX = $50;
  AsmPushEDX = $52;
  AsmJmpShort = $E9;

type
  TStub = packed record
    PopEDX: Byte;
    MovEAX: Byte;
    SelfPointer: Pointer;
    PushEAX: Byte;
    PushEDX: Byte;
    JmpShort: Byte;
    Displacement: Integer;
  end;
{$ENDIF CPUX64}

constructor TCallBackStub.Create(Obj: TObject; MethodPtr: Pointer;
  NumArgs: integer);
{$IFNDEF CPUX64}
var
  Stub: ^TStub;
begin
  // Allocate memory for the stub
  // 1/10/04 Support for 64 bit, executable code must be in virtual space
  Stub := VirtualAlloc(nil, SizeOf(TStub), MEM_COMMIT, PAGE_EXECUTE_READWRITE);

  // Pop the return address off the stack
  Stub^.PopEDX := AsmPopEDX;

  // Push the object pointer on the stack
  Stub^.MovEAX := AsmMovEAX;
  Stub^.SelfPointer := Obj;
  Stub^.PushEAX := AsmPushEAX;

  // Push the return address back on the stack
  Stub^.PushEDX := AsmPushEDX;

  // Jump to the 'real' procedure, the method.
  Stub^.JmpShort := AsmJmpShort;
  Stub^.Displacement := (Integer(MethodPtr) - Integer(@(Stub^.JmpShort))) -
    (SizeOf(Stub^.JmpShort) + SizeOf(Stub^.Displacement));

  // Return a pointer to the stub
  fCodeSize := SizeOf(TStub);
  fStubPointer := Stub;
{$ELSE CPUX64}
const
RegParamCount = 4;
ShadowParamCount = 4;

Size32Bit = 4;
Size64Bit = 8;

ShadowStack   = ShadowParamCount * Size64Bit;
SkipParamCount = RegParamCount - ShadowParamCount;

StackSrsOffset = 3;
c64stack: array[0..14] of byte = (
$48, $81, $ec, 00, 00, 00, 00,//     sub rsp,$0
$4c, $89, $8c, $24, ShadowStack, 00, 00, 00//     mov [rsp+$20],r9
);

CopySrcOffset=4;
CopyDstOffset=4;
c64copy: array[0..15] of byte = (
$4c, $8b, $8c, $24,  00, 00, 00, 00,//     mov r9,[rsp+0]
$4c, $89, $8c, $24, 00, 00, 00, 00//     mov [rsp+0],r9
);

RegMethodOffset = 10;
RegSelfOffset = 11;
c64regs: array[0..28] of byte = (
$4d, $89, $c1,      //   mov r9,r8
$49, $89, $d0,      //   mov r8,rdx
$48, $89, $ca,      //   mov rdx,rcx
$48, $b9, 00, 00, 00, 00, 00, 00, 00, 00, // mov rcx, Obj
$48, $b8, 00, 00, 00, 00, 00, 00, 00, 00 // mov rax, MethodPtr
);

c64jump: array[0..2] of byte = (
$48, $ff, $e0  // jump rax
);

CallOffset = 6;
c64call: array[0..10] of byte = (
$48, $ff, $d0,    //    call rax
$48, $81,$c4,  00, 00, 00, 00,   //     add rsp,$0
$c3// ret
);
var
  i: Integer;
  P,PP,Q: PByte;
  lCount : integer;
  lSize : integer;
  lOffset : integer;
begin
    lCount := SizeOf(c64regs);
    if NumArgs>=RegParamCount then
       Inc(lCount,sizeof(c64stack)+(NumArgs-RegParamCount)*sizeof(c64copy)+sizeof(c64call))
    else
       Inc(lCount,sizeof(c64jump));

    Q := VirtualAlloc(nil, lCount, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    P := Q;

    lSize := 0;
    if NumArgs>=RegParamCount then
    begin
        lSize := ( 1+ ((NumArgs + 1 - SkipParamCount) div 2) * 2 )* Size64Bit;   // 16 byte stack align

        pp := p;
        move(c64stack,P^,SizeOf(c64stack));
        Inc(P,StackSrsOffset);
        move(lSize,P^,Size32Bit);
        p := pp;
        Inc(P,SizeOf(c64stack));
        for I := 0 to NumArgs - RegParamCount -1 do
        begin
            pp := p;
            move(c64copy,P^,SizeOf(c64copy));
            Inc(P,CopySrcOffset);
            lOffset := lSize + (i+ShadowParamCount+1)*Size64Bit;
            move(lOffset,P^,Size32Bit);
            Inc(P,CopyDstOffset+Size32Bit);
            lOffset := (i+ShadowParamCount+1)*Size64Bit;
            move(lOffset,P^,Size32Bit);
            p := pp;
            Inc(P,SizeOf(c64copy));
        end;
    end;

    pp := p;
    move(c64regs,P^,SizeOf(c64regs));
    Inc(P,RegSelfOffset);
    move(Obj,P^,SizeOf(Obj));
    Inc(P,RegMethodOffset);
    move(MethodPtr,P^,SizeOf(MethodPtr));
    p := pp;
    Inc(P,SizeOf(c64regs));

    if NumArgs<RegParamCount then
      move(c64jump,P^,SizeOf(c64jump))
    else
    begin
      move(c64call,P^,SizeOf(c64call));
      Inc(P,CallOffset);
      move(lSize,P^,Size32Bit);
    end;
    fCodeSize := lcount;
   fStubPointer := Q;
{$ENDIF CPUX64}
end;

destructor TCallBackStub.Destroy;
begin
  VirtualFree(fStubPointer, fCodeSize, MEM_DECOMMIT);
  inherited;
end;

function TCallBackStub.GetStubPointer: Pointer;
begin
  Result := fStubPointer;
end;

initialization
  Assert(IsUnicode, 'Windows veersion not supported');
  PIDLMgr := TCommonPIDLManager.Create;
  LoadWideFunctions;

// If this unit is to be weak packages this must be removed
finalization
  FreeAndNil(PIDLMgr);
  CommonUnloadAllLibraries;
end.

