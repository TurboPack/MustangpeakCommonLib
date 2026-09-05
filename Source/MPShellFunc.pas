unit MPShellFunc;

interface

uses
  System.UITypes;

type
  TColorHelper = record helper for TColor
  public
    class function FromUInt32(const AValue: UInt32): TColor; static; inline;
    function ToUInt32: UInt32; inline;
  end;

function AsString(const AValue: AnsiString): string; inline;
function ToAnsiString(const AValue: string): AnsiString;
function ToString(const AValue: AnsiString): string;

function ToInt8(const AValue: Int16): Int8; overload; inline;
function ToInt8(const AValue: Int32): Int8; overload; inline;
function ToInt8(const AValue: Int64): Int8; overload; inline;

function ToUInt8(const AValue: Int16): UInt8; overload; inline;
function ToUInt8(const AValue: Int32): UInt8; overload; inline;
function ToUInt8(const AValue: Int64): UInt8; overload; inline;

function ToInt16(const AValue: UInt16): Int16; overload; inline;
function ToInt16(const AValue: Int32): Int16; overload; inline;
function ToInt16(const AValue: Int64): Int16; overload; inline;

function ToUInt16(const AValue: Int32): UInt16; overload; inline;
function ToUInt16(const AValue: Int64): UInt16; overload; inline;

function ToInt32(const AValue: Int32): Int32; overload; inline;
function ToInt32(const AValue: Int64): Int32; overload; inline;
function ToInt32(const AValue: UInt32): Int32; overload; inline;
function ToInt32(const AValue: UInt64): Int32; overload; inline;

function ToInt64(const AValue: Int32): Int64; overload; inline;
function ToInt64(const AValue: UInt32): Int64; overload; inline;
function ToInt64(const AValue: Int64): Int64; overload; inline;
function ToInt64(const AValue: UInt64): Int64; overload; inline;

function ToUInt32(const AValue: Int32): UInt32; overload; inline;
function ToUInt32(const AValue: Int64): UInt32; overload; inline;
function ToUInt32(const AValue: UInt32): UInt32; overload; inline;
function ToUInt32(const AValue: UInt64): UInt32; overload; inline;

function ToUInt64(const AValue: Int32): UInt64; overload; inline;
function ToUInt64(const AValue: Int64): UInt64; overload; inline;
function ToUInt64(const AValue: UInt32): UInt64; overload; inline;
function ToUInt64(const AValue: UInt64): UInt64; overload; inline;

function ToNativeInt(const AValue: Int32): NativeInt; overload; inline;
function ToNativeInt(const AValue: UInt32): NativeInt; overload; inline;
function ToNativeInt(const AValue: Int64): NativeInt; overload; inline;
function ToNativeInt(const AValue: UInt64): NativeInt; overload; inline;

function ToNativeUInt(const AValue: Int32): NativeUInt; overload; inline;
function ToNativeUInt(const AValue: UInt32): NativeUInt; overload; inline;
function ToNativeUInt(const AValue: Int64): NativeUInt; overload; inline;
function ToNativeUInt(const AValue: UInt64): NativeUInt; overload; inline;

implementation

uses
  System.SysUtils;

function AsString(const AValue: AnsiString): string;
begin
  Result := ToString(AValue);
end;

function ToAnsiString(const AValue: string): AnsiString;
var
  lBytes: TBytes;
begin
  lBytes := TEncoding.ANSI.GetBytes(AValue);
  if Assigned(lBytes) then
  begin
    SetLength(Result, Length(lBytes));
    Move(lBytes[0], Result[1], Length(lBytes));
  end
  else
    Result := '';
end;

function ToString(const AValue: AnsiString): string;
var
  lBytes: TBytes;
begin
  if AValue = string.Empty then
    Exit(string.Empty);

  SetLength(lBytes, Length(AValue));
  Move(AValue[1], lBytes[0], Length(lBytes));

  Result := TEncoding.ANSI.GetString(lBytes);
end;

function ToInt8(const AValue: Int16): Int8;
begin
  Result := Int8(AValue);
end;

function ToInt8(const AValue: Int32): Int8;
begin
  Result := Int8(AValue);
end;

function ToInt8(const AValue: Int64): Int8;
begin
  Result := Int8(AValue);
end;

function ToUInt8(const AValue: Int16): UInt8;
begin
  Result := UInt8(AValue);
end;

function ToUInt8(const AValue: Int32): UInt8;
begin
  Result := UInt8(AValue);
end;

function ToUInt8(const AValue: Int64): UInt8;
begin
  Result := UInt8(AValue);
end;

function ToInt16(const AValue: UInt16): Int16;
begin
  Result := Int16(AValue);
end;

function ToInt16(const AValue: Int32): Int16;
begin
  Result := Int16(AValue);
end;

function ToInt16(const AValue: Int64): Int16;
begin
  Result := Int16(AValue);
end;

function ToUInt16(const AValue: Int32): UInt16;
begin
  Result := UInt16(AValue);
end;

function ToUInt16(const AValue: Int64): UInt16;
begin
  Result := UInt16(AValue);
end;

function ToInt32(const AValue: Int32): Int32;
begin
  Result := Int32(AValue);
end;

function ToInt32(const AValue: Int64): Int32;
begin
  Result := Int32(AValue);
end;

function ToInt32(const AValue: UInt32): Int32;
begin
  Result := Int32(AValue);
end;

function ToInt32(const AValue: UInt64): Int32;
begin
  Result := Int32(AValue);
end;

function ToInt64(const AValue: Int32): Int64; overload; inline;
begin
  Result := Int64(AValue);
end;

function ToInt64(const AValue: UInt32): Int64; overload; inline;
begin
  Result := Int64(AValue);
end;

function ToInt64(const AValue: Int64): Int64; overload; inline;
begin
  Result := Int64(AValue);
end;

function ToInt64(const AValue: UInt64): Int64; overload; inline;
begin
  Result := Int64(AValue);
end;

function ToUInt32(const AValue: Int32): UInt32;
begin
  Result := UInt32(AValue);
end;

function ToUInt32(const AValue: Int64): UInt32;
begin
  Result := UInt32(AValue);
end;

function ToUInt32(const AValue: UInt32): UInt32;
begin
  Result := UInt32(AValue);
end;

function ToUInt32(const AValue: UInt64): UInt32;
begin
  Result := UInt32(AValue);
end;

function ToUInt64(const AValue: Int32): UInt64;
begin
  Result := UInt64(AValue);
end;

function ToUInt64(const AValue: Int64): UInt64;
begin
  Result := UInt64(AValue);
end;

function ToUInt64(const AValue: UInt32): UInt64;
begin
  Result := UInt64(AValue);
end;

function ToUInt64(const AValue: UInt64): UInt64;
begin
  Result := UInt64(AValue);
end;

function ToNativeInt(const AValue: Int32): NativeInt; overload; inline;
begin
  Result := NativeInt(AValue);
end;

function ToNativeInt(const AValue: UInt32): NativeInt; overload; inline;
begin
  Result := NativeInt(AValue);
end;

function ToNativeInt(const AValue: Int64): NativeInt; overload; inline;
begin
  Result := NativeInt(AValue);
end;

function ToNativeInt(const AValue: UInt64): NativeInt; overload; inline;
begin
  Result := NativeInt(AValue);
end;

function ToNativeUInt(const AValue: Int32): NativeUInt; overload; inline;
begin
  Result := NativeUInt(AValue);
end;

function ToNativeUInt(const AValue: UInt32): NativeUInt; overload; inline;
begin
  Result := NativeUInt(AValue);
end;

function ToNativeUInt(const AValue: Int64): NativeUInt; overload; inline;
begin
  Result := NativeUInt(AValue);
end;

function ToNativeUInt(const AValue: UInt64): NativeUInt; overload; inline;
begin
  Result := NativeUInt(AValue);
end;

{ TColorHelper }

class function TColorHelper.FromUInt32(const AValue: UInt32): TColor;
begin
  Result := TColor(AValue);
end;

function TColorHelper.ToUInt32: UInt32;
begin
  Result := UInt32(Self);
end;

end.
