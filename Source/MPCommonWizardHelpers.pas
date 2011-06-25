unit MPCommonWizardHelpers;

//
// This unit is to be used in DesignTime packages and units ONLY
//

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
//----------------------------------------------------------------------------

interface

{$include Compilers.Inc}
{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

uses
  SysUtils,
  Windows,
  Controls,
  ImgList,
  ToolsApi,
  Dialogs,
  Menus,
  {$IFDEF COMPILER_6_UP}
  DesignIntf,
  DesignEditors,
  TreeIntf,
  VCLEditors,
  TypInfo,
  Graphics,
  {$ELSE}
  DsgnIntf,
  DsgnWnds,
    {$IFDEF COMPILER_6_UP}
    ParentageSupport, // Sprig support in DsnIDE5.dpk, D4 has no equivalent
    {$ELSE}
      {$IFDEF DELPHI_5}
      ParentageSupport, // Sprig support in DsnIDE5.dpk, D4 has no equivalent
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
  Classes,
  MPCommonWizardTemplates;

const
  DELPHI_USES_UNITS: array[0..7] of string = (
  'Windows',
  'Messages',
  'SysUtils',
  'Classes',
  'Controls',
  'Forms',
  'Dialogs',
  'Graphics'
  );

  BUIDLER_INCLUDE: array[0..0] of string = (
  'vcl.h'
  );

type
  {$IFNDEF COMPILER_7_UP}
  // ***************************************************************************
  // TOTAFile
  //   The implements the basic IOTAFile interface
  // ***************************************************************************
  TOTAFile = class(TInterfacedObject, IOTAFile)
    FSource: string;
    FAge: TDateTime;
  public
    constructor Create(const StringCode: String; const Age: TDateTime = -1);
    function GetSource: string; virtual;
    function GetAge: TDateTime; virtual;
  end;
  {$ENDIF}

  {$IFNDEF COMPILER_5_UP}
  TNotifierObject = class(TInterfacedObject)
  public
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
  end;
  {$ENDIF}

  // ---------------------------------------------------------------------------
  // OTAFile Module
  // ---------------------------------------------------------------------------
  TCommonOTAFile = class(TOTAFile, IOTAFile)
  private
    FAncestorIdent: string;
    FFormIdent: string;
    FModuleIdent: string;
    FUsesIdent: TStringList;
  public
    property AncestorIdent: string read FAncestorIdent write FAncestorIdent;
    property FormIdent: string read FFormIdent write FFormIdent;
    property ModuleIdent: string read FModuleIdent write FModuleIdent;
    property IncludeIdent: TStringList read FUsesIdent write FUsesIdent;
  end;

  TCommonOTAFileForm = class(TCommonOTAFile, IOTAFile)
  private
  public
    function GetSource: string; override;
  end;

  // ---------------------------------------------------------------------------
  // ModuleCreator Classes
  // ---------------------------------------------------------------------------

  // ***************************************************************************
  // TCommonWizardModuleCreator
  //   Implements the basic functionality of IOTAModuleCreator
  // ***************************************************************************
  TCommonWizardModuleCreator = class(TInterfacedObject, IOTACreator, IOTAModuleCreator)
  private
    FAncestorName: string;
    FFormName: string;
    FMainForm: Boolean;
    FShowForm: Boolean;
    FShowSource: Boolean;
    FUsesIdent: TStringList;
    function GetIsDelphi: Boolean;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    // Override to initalize the creator
    procedure InitializeCreator; virtual;

    // IOTACreator
    function GetCreatorType: string; virtual;
    function GetExisting: Boolean; virtual;
    function GetFileSystem: string; virtual;
    function GetOwner: IOTAModule; virtual;
    function GetUnnamed: Boolean; virtual;

    // IOTAModuleCreator
    function GetAncestorName: string; virtual;
    function GetImplFileName: string; virtual;
    function GetIntfFileName: string;virtual;
    function GetFormName: string; virtual;
    function GetMainForm: Boolean; virtual;
    function GetShowForm: Boolean; virtual;
    function GetShowSource: Boolean; virtual;
    function NewFormFile(const FormIdent, AncestorIdent: string): IOTAFile; virtual;
    function NewImplSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile; virtual;
    function NewIntfSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile; virtual;
    procedure FormCreated(const FormEditor: IOTAFormEditor); virtual;

    procedure LoadDefaultBuilderIncludeStrings(ClearFirst: Boolean);
    procedure LoadDefaultDelphiUsesStrings(ClearFirst: Boolean);
    
    property CreatorType: string read GetCreatorType;
    property Existing: Boolean read GetExisting;
    property FileSystem: string read GetFileSystem;
    property Owner: IOTAModule read GetOwner;
    property Unnamed: Boolean read GetUnnamed;

    property AncestorName: string read GetAncestorName write FAncestorName;
    property FormName: string read GetFormName write FFormName;
    property ImplFileName: string read GetImplFileName;
    property IntfFileName: string read GetIntfFileName;
    property IsDelphi: Boolean read GetIsDelphi;
    property MainForm: Boolean read GetMainForm write FMainForm;
    property ShowForm: Boolean read GetShowForm write FShowForm;
    property ShowSource: Boolean read GetShowSource write FShowSource;
    property IncludeIdent: TStringList read FUsesIdent write FUsesIdent;
  end;
  TCommonWizardModuleCreatorClass = class of TCommonWizardModuleCreator;

  // ***************************************************************************
  // TCommonWizardEmptyUnitCreator
  //   The Creator that creates a basic unit for the project
  // ***************************************************************************
  TCommonWizardEmptyUnitCreator = class(TCommonWizardModuleCreator)
  public
    function GetCreatorType: string; override;
  end;

  // ***************************************************************************
  // TCommonWizardEmptyFormCreator
  //   The Creator that creates a basic blank Form for the project
  // ***************************************************************************
  TCommonWizardEmptyFormCreator = class(TCommonWizardModuleCreator)
  private
  public
    function GetCreatorType: string; override;
    function NewImplSource(const ModuleIdent: string; const FormIdent: string; const AncestorIdent: string): IOTAFile; override;
  end;

  // ***************************************************************************
  // TCommonWizardEmptyTextCreator
  //   The Creator that creates a basic text file for the project
  // ***************************************************************************
  TCommonWizardEmptyTextCreator = class(TCommonWizardModuleCreator)
  public
    function GetCreatorType: string; override;
  end;

  // ---------------------------------------------------------------------------
  // Repository Wizard Classes
  // ---------------------------------------------------------------------------

  // ***************************************************************************
  // TCommonWizardNotifierObject
  //
  //  Use as basis for Wizards.
  //
  // ***************************************************************************
  TCommonWizardNotifierObject = class(TNotifierObject,
    {$IFDEF COMPILER_6_UP}IOTARepositoryWizard60, {$ENDIF}
    {$IFDEF COMPILER_8_UP}IOTARepositoryWizard80, {$ENDIF}
    IOTAWizard,
    IOTARepositoryWizard,
    IOTAProjectWizard)
  private
    FAuthor: string;
    FCaption: string;
    FComment: string;
    FGlyphResourceID: string;
    FPage: string;
    FState: TWizardState;
    FUniqueID: string;
    {$IFDEF COMPILER_8_UP}
    FGalleryCategory: IOTAGalleryCategory;
    {$ENDIF COMPILER_8_UP}
  protected
    function GetGlpyhResourceID: string; virtual;
  public
    constructor Create;

    // Override to load the wizard with the necessary information
    procedure InitializeWizard; virtual; 

    // IOTAWizard
    function GetIDString: string; virtual;
    function GetName: string; virtual;
    function GetState: TWizardState;
    procedure Execute; virtual; 
    //
    function GetAuthor: string;
    function GetComment: string; virtual;
    function GetPage: string; virtual;
    {$IFDEF COMPILER_6_UP}
    function GetGlyph: Cardinal;
    {$ELSE}
    function GetGlyph: HICON;
    {$ENDIF}

    {$IFDEF COMPILER_6_UP}
    // IOTARepositoryWizard60
    function GetDesigner: string;
    {$ENDIF}

    {$IFDEF COMPILER_8_UP}
    // IOTARepositoryWizard80
    function GetGalleryCategory: IOTAGalleryCategory; virtual;
    function GetPersonality: string; virtual; abstract;
    {$ENDIF}

    {$IFDEF COMPILER_6_UP}
    property Designer: string read GetDesigner;
    {$ENDIF}
    {$IFDEF COMPILER_8_UP}
    property Personality: string read GetPersonality;
    {$ENDIF}

    // Set to the Author that will show up in the Details View of the Object
    property Author: string read GetAuthor write FAuthor;
    // Set to the Caption that will show up in the View of the Object
    property Caption: string read GetName write FCaption;
    // Set to the Comment that will show up in the Details View of the Object
    property Comment: string read GetComment write FComment;
    // Set to the Resource ID (in the a *.res file) that defines the icon for the object
    {$IFDEF COMPILER_8_UP}
    // Set the Gallery Category where the Wizard reside, it is the value returned from AddDelphiCategory or AddBuilderCategory
    property GalleryCategory: IOTAGalleryCategory read GetGalleryCategory write FGalleryCategory;
    {$ENDIF COMPILER_8_UP}
    property GlyphResourceID: string read GetIDString write FGlyphResourceID;
    // Set to the Page that Object will show up on when selecting File>New>Other...
    property Page: string read GetPage write FPage;
    // Set special attributes for the Wizard (enabled, checked etc)
    property State: TWizardState read GetState write FState;
    // Set to a unique string that will identify the Object to Delphi (i.e. "Mustangpeak.CommonWizards.Demo")
    property UniqueID: string read GetIDString write FUniqueID;  
  end;

   TCommonWizardModuleCreate = class(TCommonWizardNotifierObject,
    {$IFDEF COMPILER_6_UP}IOTARepositoryWizard60, {$ENDIF}
    {$IFDEF COMPILER_8_UP}IOTARepositoryWizard80, {$ENDIF}
    IOTAWizard,
    IOTARepositoryWizard,
    IOTAProjectWizard)
  private
    FCreatorClass: TCommonWizardModuleCreatorClass;
  public
    procedure Execute; override;

    property CreatorClass: TCommonWizardModuleCreatorClass read FCreatorClass write FCreatorClass;
   end;

  // ***************************************************************************
  // TCommonWizardDelphiForm
  //   Wizard to Create a new Delphi Form
   // ***************************************************************************
  TCommonWizardDelphiForm = class(TCommonWizardModuleCreate,
    {$IFDEF COMPILER_6_UP}IOTARepositoryWizard60, {$ENDIF}
    {$IFDEF COMPILER_8_UP}IOTARepositoryWizard80, {$ENDIF}
    IOTAWizard,
    IOTARepositoryWizard,
    IOTAProjectWizard)
  protected
  public
    // IOTAWizard
    {$IFDEF COMPILER_8_UP}
    function GetPersonality: string; override;
    {$ENDIF COMPILER_8_UP}
  end;

  TCommonWizardBuilderForm = class(TCommonWizardDelphiForm,
    {$IFDEF COMPILER_6_UP}IOTARepositoryWizard60, {$ENDIF}
    {$IFDEF COMPILER_8_UP}IOTARepositoryWizard80, {$ENDIF}
    IOTAWizard,
    IOTARepositoryWizard,
    IOTAProjectWizard)
  public
    // IOTAWizard
    {$IFDEF COMPILER_8_UP}
    function GetPersonality: string; override;
    {$ENDIF COMPILER_8_UP}
  end;

  {$IFDEF COMPILER_6_UP}
  TPersistentHack = class(TPersistent);

  // ***************************************************************************
  // TImageIndexProperty
  //   Creates a property that implmements a customdraw dropdown list for
  // ImageList index properties.
  // ***************************************************************************
  TCommonImageIndexProperty = class(TIntegerProperty, ICustomPropertyDrawing, ICustomPropertyListDrawing)
  private
    function GetImageList: TCustomImageList;
    function GetImageListAt(ComponentIndex: Integer): TCustomImageList;
  protected
    function ExtractImageList(Inst: TPersistent; out ImageList: TCustomImageList): Boolean; virtual;
    property ImageList: TCustomImageList read GetImageList;
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
    function GetValue: string; override;
    procedure SetValue(const Value: string); override;

    procedure ListMeasureWidth(const Value: string; ACanvas: TCanvas; var AWidth: Integer);
    procedure ListMeasureHeight(const Value: string; ACanvas: TCanvas; var AHeight: Integer);
    procedure ListDrawValue(const Value: string; ACanvas: TCanvas; const ARect: TRect; ASelected: Boolean);

    procedure PropDrawName(ACanvas: TCanvas; const ARect: TRect; ASelected: Boolean);
    procedure PropDrawValue(ACanvas: TCanvas; const ARect: TRect; ASelected: Boolean);
  end;
{$ENDIF COMPILER_6_UP}


function GetCurrentProjectGroup: IOTAProjectGroup;
function GetCurrentProject: IOTAProject;
{$IFDEF COMPILER_8_UP}
// These must be called in the initialization section of a unit
function AddDelphiCategory(CategoryID, CategoryCaption: string): IOTAGalleryCategory;
function AddBuilderCategory(CategoryID, CategoryCaption: string): IOTAGalleryCategory;
procedure RemoveCategory(Category: IOTAGalleryCategory);

function IsDelphiPersonality: Boolean;
{$ENDIF COMPILER_8_UP}

implementation


function GetCurrentProjectGroup: IOTAProjectGroup;
var
  IModuleServices: IOTAModuleServices;
  IModule: IOTAModule;
  IProjectGroup: IOTAProjectGroup;
  i: Integer;
begin
  Result := nil;
  IModuleServices := BorlandIDEServices as IOTAModuleServices;
  for i := 0 to IModuleServices.ModuleCount - 1 do
  begin
    IModule := IModuleServices.Modules[i];
    if IModule.QueryInterface(IOTAProjectGroup, IProjectGroup) = S_OK then
    begin
      Result := IProjectGroup;
      Break;
    end;
  end;
end;

function GetCurrentProject: IOTAProject;
var
  ProjectGroup: IOTAProjectGroup;
 // i: Integer;
begin
  Result := nil;
  ProjectGroup := GetCurrentProjectGroup;

  if Assigned(ProjectGroup) then
    if ProjectGroup.ProjectCount > 0 then
      Result := ProjectGroup.ActiveProject
end;

{$IFDEF COMPILER_8_UP}
function AddDelphiCategory(CategoryID, CategoryCaption: string): IOTAGalleryCategory;
var
  Manager: IOTAGalleryCategoryManager;
  ParentCategory: IOTAGalleryCategory;
begin
  Result := nil;
  Manager := BorlandIDEServices as IOTAGalleryCategoryManager;
  if Assigned(Manager) then
  begin
    ParentCategory := Manager.FindCategory(sCategoryDelphiNew);
    if Assigned(ParentCategory) then
      Result := Manager.AddCategory(ParentCategory, CategoryID, CategoryCaption);
  end
end;

function AddBuilderCategory(CategoryID, CategoryCaption: string): IOTAGalleryCategory;
var
  Manager: IOTAGalleryCategoryManager;
  ParentCategory: IOTAGalleryCategory;
begin
  Result := nil;
  Manager := BorlandIDEServices as IOTAGalleryCategoryManager;
  if Assigned(Manager) then
  begin
    ParentCategory := Manager.FindCategory(sCategoryCBuilderNew);
    if Assigned(ParentCategory) then
      Result := Manager.AddCategory(ParentCategory, CategoryID, CategoryCaption);
  end
end;

procedure RemoveCategory(Category: IOTAGalleryCategory);
var
  Manager: IOTAGalleryCategoryManager;
begin
  Manager := BorlandIDEServices as IOTAGalleryCategoryManager;
  if Assigned(Manager) then
  begin
    if Assigned(Category) then
      Manager.DeleteCategory(Category)
  end
end;

function IsDelphiPersonality: Boolean;
var
  Personalities: IOTAPersonalityServices;
begin
  Personalities := BorlandIDEServices as IOTAPersonalityServices;
  Result := Personalities.CurrentPersonality = sDelphiPersonality;
end;
{$ENDIF}

{$IFNDEF COMPILER_7_UP}
{ TOTAFile }

constructor TOTAFile.Create(const StringCode: String; const Age: TDateTime);
begin
  FSource := StringCode;
  FAge := Age;
end;

function TOTAFile.GetSource: string;
begin
  Result := FSource;
end;

function TOTAFile.GetAge: TDateTime;
begin
  Result := FAge;
end;
{$ENDIF}


{$IFNDEF COMPILER_5_UP}
{ TNotifierObject }

procedure TNotifierObject.AfterSave;
begin
  // do nothing stub implementation
end;

procedure TNotifierObject.BeforeSave;
begin
  // do nothing stub implementation
end;

procedure TNotifierObject.Destroyed;
begin
  // do nothing stub implementation
end;

procedure TNotifierObject.Modified;
begin
  // do nothing stub implementation
end;
{$ENDIF}

{ TCommonWizardDelphiForm }

{$IFDEF COMPILER_8_UP}
function TCommonWizardDelphiForm.GetPersonality: string;
begin
  Result := sDelphiPersonality
end;
{$ENDIF COMPILER_8_UP}

{ TCommonWizardBuilderForm }

{$IFDEF COMPILER_8_UP}
function TCommonWizardBuilderForm.GetPersonality: string;
begin
  Result := sCBuilderPersonality
end;
{$ENDIF COMPILER_8_UP}

{$IFDEF COMPILER_6_UP}

{ TImageIndexProperty}

function TCommonImageIndexProperty.ExtractImageList(Inst: TPersistent; out ImageList: TCustomImageList): Boolean;
var
  P: PPropList;
  I, C: Integer;
  s: string;
begin
  s := Inst.ClassName;
  C := GetPropList(Inst.ClassInfo, P);
  try
    for I := 0 to C - 1 do
      if (P[I].PropType^.Kind = tkClass) and GetTypeData(P[I].PropType^).ClassType.InheritsFrom(TCustomImageList) then
      begin
        Result := True;
        ImageList := TCustomImageList(TypInfo.GetObjectProp(Inst, P[I]));
        Exit;
      end;
    Result := False;
  finally
    FreeMem(P);
  end;
end;


function TCommonImageIndexProperty.GetImageListAt(ComponentIndex: Integer): TCustomImageList;
var
  Inst: TPersistent;
begin
  Inst := GetComponent(ComponentIndex);
  while Assigned(Inst) do
  begin
    if ExtractImageList(Inst, Result) then Exit;

    Inst := TPersistentHack(Inst).GetOwner;
  end;
  Result := nil;
end;


function TCommonImageIndexProperty.GetImageList: TCustomImageList;
var
  I, J: Integer;
  ImgList: TCustomImageList;
begin
  Result := nil;
  for I := 0 to PropCount - 1 do
  begin
    ImgList := GetImageListAt(I);
    if Assigned(ImgList) then
    begin
      for J := I + 1 to PropCount - 1 do
        if GetImageListAt(J) <> ImgList then
        begin
          Result := nil;
          Exit;
        end;
      Result := ImgList;
      Exit;
    end;
  end;
end;


function TCommonImageIndexProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paRevertable, paMultiSelect];
end;


function TCommonImageIndexProperty.GetValue: string;
begin
  Result:= IntToStr(GetOrdValue);
end;


procedure TCommonImageIndexProperty.SetValue(const Value: string);
var
  XValue: integer;
begin
  try
    XValue := StrToInt(Value);
    SetOrdValue(XValue);
  except
    inherited SetValue(Value);
  end;
end;


procedure TCommonImageIndexProperty.GetValues(Proc: TGetStrProc);
var
  XImageList: TCustomImageList;
  I: integer;
begin
  XImageList:=GetImageList;
  if Assigned(XImageList) then
    for I := 0 to XImageList.Count - 1 do
      Proc(IntToStr(i));
end;


procedure TCommonImageIndexProperty.ListMeasureWidth(const Value: string; ACanvas: TCanvas;
  var AWidth: Integer);
begin
  AWidth := AWidth + ACanvas.TextHeight('M');
  if AWidth < 17 then AWidth := 17;
end;


procedure TCommonImageIndexProperty.ListMeasureHeight(const Value: string; ACanvas: TCanvas;
  var AHeight: Integer);
var
  ImageList: TCustomImageList;
begin
  ImageList := GetImageList;
  if Assigned(ImageList) then
    AHeight := ImageList.Height + 4
  else
    AHeight := 20;
  if AHeight < 17 then AHeight := 17;
end;


procedure TCommonImageIndexProperty.ListDrawValue(const Value: string; ACanvas: TCanvas;
  const ARect: TRect; ASelected: Boolean);
var
  ImageList: TCustomImageList;
  XRight: Integer;
  XOldPenColor, XOldBrushColor: TColor;
  Index: TImageIndex;
begin
  ImageList := GetImageList;
  Index := StrToIntDef(Value, -1);
  XRight := ARect.Left;
  try
    if Assigned(ImageList) and (Index >= 0) then
    begin
      XRight := ARect.Left + ImageList.Width + 4;
      XOldPenColor := ACanvas.Pen.Color;
      XOldBrushColor := ACanvas.Brush.Color;

      ACanvas.Pen.Color := ACanvas.Brush.Color;
      ACanvas.Rectangle(ARect.Left, ARect.Top, XRight, ARect.Bottom);

      ImageList.DrawOverlay(ACanvas, ARect.Left + 2, ARect.Top + 2, Index, 0);
      ACanvas.Brush.Color := XOldBrushColor;
      ACanvas.Pen.Color := XOldPenColor;
      end;
  finally
    DefaultPropertyListDrawValue(Value, ACanvas, Rect(XRight, ARect.Top, ARect.Right, ARect.Bottom), ASelected);
  end;
end;


procedure TCommonImageIndexProperty.PropDrawName(ACanvas: TCanvas; const ARect: TRect;
  ASelected: Boolean);
begin
  DefaultPropertyDrawName(Self, ACanvas, ARect);
end;


procedure TCommonImageIndexProperty.PropDrawValue(ACanvas: TCanvas; const ARect: TRect;
  ASelected: Boolean);
var
  ImageList: TCustomImageList;
begin
  ImageList := GetImageList;
  if (GetVisualValue <> '') and Assigned(ImageList) and (ImageList.Height < 17) then
    ListDrawValue(GetVisualValue, ACanvas, ARect, True{ASelected})
  else
    DefaultPropertyDrawValue(Self, ACanvas, ARect);
end;

{$ENDIF COMPILER_6_UP}


{ TCommonWizardNotifierObject }

constructor TCommonWizardNotifierObject.Create;
begin
  inherited Create;
  InitializeWizard;
end;

function TCommonWizardNotifierObject.GetAuthor: string;
begin
  Result := FAuthor
end;

function TCommonWizardNotifierObject.GetComment: string;
begin
 Result := FComment
end;

{$IFDEF COMPILER_8_UP}
function TCommonWizardNotifierObject.GetGalleryCategory: IOTAGalleryCategory;
begin
  Result := FGalleryCategory
end;
{$ENDIF COMPILER_8_UP}

function TCommonWizardNotifierObject.GetGlpyhResourceID: string;
begin
  Result := FGlyphResourceID
end;

function TCommonWizardNotifierObject.GetIDString: string;
begin
  Result := FUniqueID
end;

function TCommonWizardNotifierObject.GetName: string;
begin
   Result := FCaption
end;

function TCommonWizardNotifierObject.GetPage: string;
begin
  Result := FPage
end;

function TCommonWizardNotifierObject.GetState: TWizardState;
begin
  Result := FState
end;

{$IFDEF COMPILER_6_UP}
function TCommonWizardNotifierObject.GetGlyph: Cardinal;
begin
  Result := LoadIconA(hInstance, PAnsiChar( AnsiString(GetGlpyhResourceID)));
end;
{$ELSE}
function TCommonWizardNotifierObject.GetGlyph: HICON;
begin
  Result := LoadIcon(hInstance, PAnsiChar(GetGlpyhResourceID));
end;
{$ENDIF}

{$IFDEF COMPILER_6_UP}

function TCommonWizardNotifierObject.GetDesigner: string;
begin
  Result := dVCL
end;
{$ENDIF}

procedure TCommonWizardNotifierObject.InitializeWizard;
begin
  // Override in descendent
end;

procedure TCommonWizardNotifierObject.Execute;
begin
  // Override in descendent
end;

{ TCommonWizardModuleCreator }

constructor TCommonWizardModuleCreator.Create;
begin
  inherited;
  ShowSource := True;
  ShowForm := True;
  IncludeIdent := TStringList.Create;
  if IsDelphi then
    LoadDefaultDelphiUsesStrings(True)
  else
    LoadDefaultBuilderIncludeStrings(True);
  InitializeCreator
end;

destructor TCommonWizardModuleCreator.Destroy;
begin
  inherited;
  IncludeIdent.Free
end;

function TCommonWizardModuleCreator.GetIsDelphi: Boolean;
begin
{$IFDEF CPPB}
  Result := False;
{$ELSE}
  {$IFDEF COMPILER_10_UP}
    if IsDelphiPersonality then
      Result := True
    else
      Result := False;
  {$ELSE}
    Result := True;
  {$ENDIF}
{$ENDIF}
end;

procedure TCommonWizardModuleCreator.FormCreated(const FormEditor: IOTAFormEditor);
begin

end;

function TCommonWizardModuleCreator.GetAncestorName: string;
begin
  Result := FAncestorName
end;

function TCommonWizardModuleCreator.GetCreatorType: string;
begin
  Result := ''
end;

function TCommonWizardModuleCreator.GetExisting: Boolean;
begin
  Result := False;
end;

function TCommonWizardModuleCreator.GetFileSystem: string;
begin
  Result := '';
end;

function TCommonWizardModuleCreator.GetFormName: string;
begin
  Result := FFormName
end;

function TCommonWizardModuleCreator.GetImplFileName: string;
begin
  Result := ''
end;

function TCommonWizardModuleCreator.GetIntfFileName: string;
begin
  Result := ''
end;

function TCommonWizardModuleCreator.GetMainForm: Boolean;
begin
  Result := FMainForm
end;

function TCommonWizardModuleCreator.GetOwner: IOTAModule;
begin
  Result := GetCurrentProjectGroup;
  if Assigned(Result) then
    Result := (Result as IOTAProjectGroup).ActiveProject
  else
    Result := GetCurrentProject
end;

function TCommonWizardModuleCreator.GetShowForm: Boolean;
begin
  Result := FShowForm
end;

function TCommonWizardModuleCreator.GetShowSource: Boolean;
begin
  Result := FShowSource
end;

function TCommonWizardModuleCreator.GetUnnamed: Boolean;
begin
  Result := True
end;

function TCommonWizardModuleCreator.NewFormFile(const FormIdent,
  AncestorIdent: string): IOTAFile;
begin
  Result := nil
end;

function TCommonWizardModuleCreator.NewImplSource(const ModuleIdent,
  FormIdent, AncestorIdent: string): IOTAFile;
begin
  Result := nil
end;

function TCommonWizardModuleCreator.NewIntfSource(const ModuleIdent,
  FormIdent, AncestorIdent: string): IOTAFile;
begin
  Result := nil
end;

procedure TCommonWizardModuleCreator.LoadDefaultBuilderIncludeStrings(ClearFirst: Boolean);
var
  i: Integer;
begin
  if ClearFirst then
    IncludeIdent.Clear;
  for i := 0 to High(BUIDLER_INCLUDE) do
    IncludeIdent.Add(BUIDLER_INCLUDE[i])
end;

procedure TCommonWizardModuleCreator.LoadDefaultDelphiUsesStrings(ClearFirst: Boolean);
var
  i: Integer;
begin
  if ClearFirst then
    IncludeIdent.Clear;
  for i := 0 to High(DELPHI_USES_UNITS) do
    IncludeIdent.Add(DELPHI_USES_UNITS[i])
end;

procedure TCommonWizardModuleCreator.InitializeCreator;
begin
  // Override in descendent
end;

{ TCommonWizardEmptyUnitCreator }

function TCommonWizardEmptyUnitCreator.GetCreatorType: string;
begin
  Result := sUnit;
end;

{ TCommonWizardEmptyFormCreator }

function TCommonWizardEmptyFormCreator.GetCreatorType: string;
begin
  Result := sForm
end;

function TCommonWizardEmptyFormCreator.NewImplSource(const ModuleIdent: string; const FormIdent: string; const AncestorIdent: string): IOTAFile;
var
  OTAFile: TCommonOTAFileForm;
begin
  Result := nil;
  // Create the default source code for a new application
  // Slip in the default ProjectName to the IOTAFile instance
  OTAFile := TCommonOTAFileForm.Create(FormIdent);
  OTAFile.ModuleIdent := ModuleIdent;
  OTAFile.FormIdent := FormIdent;
  OTAFile.AncestorIdent := AncestorIdent;    
  OTAFile.IncludeIdent := IncludeIdent;
  Result := OTAFile as IOTAFile;
end;

{ TCommonWizardEmptyTextCreator }

function TCommonWizardEmptyTextCreator.GetCreatorType: string;
begin
  Result := sText
end;   

{ TCommonWizardModuleCreate }

procedure TCommonWizardModuleCreate.Execute;
var
  Module: IOTAModule;
begin
  if Assigned(CreatorClass) then
    Module := (BorlandIDEServices as IOTAModuleServices).CreateModule(CreatorClass.Create)
  else
    beep(500, 50);
end;

{ TCommonOTAFileForm }
function TCommonOTAFileForm.GetSource: string;
var
  UsesClause: string;
  i: Integer;
  IsBCB: Boolean;
begin
{$IFDEF CPPB}
  Result := FILE_FORM_TEMPLATE_BCB;
  IsBCB := True;
{$ELSE}
  {$IFDEF COMPILER_10_UP}
    if IsDelphiPersonality then
    begin
      Result := FILE_FORM_TEMPLATE_DELPHI;
      IsBCB := False
    end else
    begin
      Result := FILE_FORM_TEMPLATE_BCB;
      IsBCB := True
    end;
  {$ELSE}
    Result := FILE_FORM_TEMPLATE_DELPHI;
    IsBCB := False;
  {$ENDIF}
{$ENDIF}
  Result := StringReplace(Result, '%FormIdent', FormIdent, [rfIgnoreCase, rfReplaceAll]);
  Result := StringReplace(Result, '%AncestorIdent', AncestorIdent, [rfIgnoreCase, rfReplaceAll]);
  Result := StringReplace(Result, '%ModuleIdent', ModuleIdent, [rfIgnoreCase, rfReplaceAll]);
  UsesClause := '';
  for i := 0 to IncludeIdent.Count - 1 do
  begin
    if i < IncludeIdent.Count - 1 then
    begin
      if IsBCB then
        UsesClause := UsesClause + '#include "' + IncludeIdent[i] +'"' + #13#10
      else begin
        if i = 0 then
          UsesClause := UsesClause + IncludeIdent[i] + ',' + #13#10
        else
          UsesClause := UsesClause + '  ' + IncludeIdent[i] + ',' + #13#10
      end  
    end else
    begin
      if IsBCB then
        UsesClause := UsesClause + '#include "' + IncludeIdent[i] +'"' + #13#10
      else
        UsesClause := UsesClause + '  ' + IncludeIdent[i] + ';' + #13#10
    end
  end;

  Result := StringReplace(Result, '%IncludeList', UsesClause, [rfIgnoreCase, rfReplaceAll]);
end;

end.
