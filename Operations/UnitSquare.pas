unit UnitSquare;

interface

uses System.Classes, System.SysUtils, Comserv, ComObj, Activex, MyInterface;


type


  TSquareExtraction = class(TComObject, IParameterizedOperations)
  private
  { Private declarations }
   // ����� �� �������� ����� ����������� ������
   poDataSquare : TRecordingParameters;

  protected
  { Protected declarations }

  public

//   ������������� � ��������������� ��������
    procedure Initialize;override;
    destructor Destroy;override;
//
//  IParameterizedOperations
//
  // ����������� ������� ����������������� ��������
  // ��������� ������� ����������������� �������� (����������: <>0 ������)
    function  poCheck:longint;virtual;StdCall;
  // ��������� ������� ����������������� �������� (����������: <>0 ������)
    function  poExecute:longint;virtual;StdCall;
  // �������� ��������� ������� ����������������� ��������
    procedure  poClearParameters;virtual;StdCall;

  // ����� �������� ����������������� ��������
    function poGetNameParameters:AnsiString;virtual;Stdcall;
  // ��������� ������� ����������������� �������� (����������: < 0 ������)
    function poCountOfParameters:longint;virtual;StdCall;
// ������� ����� ��������� ������ poExecute
    function poGetResult: Variant;StdCall;
///////////////////////////////////////////////////////////////////////////////////////////////////
///  IRecordingParameters
  // �������� ��������� Index ���������� � 1
    function pGetNameParameter(Index:longint) : AnsiString;StdCall;
  // ����������� ���������� ��������� Index ���������� � 1
    function pGetHintParameter(Index:longint) : AnsiString;StdCall;
  // ��� ���������(������������/��������������) Index ���������� � 1
    function pGetRequired(Index:longint) : boolean;StdCall;
  // ��� ��������� ������������ �������� VarType(������������� �����, ������� �������� �� ������� � �.�) Index ���������� � 1
    function pGetEnumeration(Index:longint) : Variant;StdCall;
  // �������� �������� Index ���������� � 1
    function pGetParameters(Index:longint)  : Variant;StdCall;
  // �������� �������� ���������� <> 0, ���� ������ Index ���������� � 1
    function pSetParameters(Index:longint; pParam : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;
  // ��������� Index ���������� � 1
    function pCheckParameters(Index:longint; pVerify : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;

  end;

//////////////////////////////////////////////////////////////
// ����������� ����������
//Function DllRegisterServer:HResult; stdcall;

Exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

implementation

uses System.Variants, System.DateUtils;

(*
Function DllRegisterServer:HResult; stdcall;
// ����������� ���������� � ��������� �����������
//////////////////////////////////////////////////////////////
// ����������� ����������
//  Function DllRegisterServer:HResult; stdcall;
// ��� ��������
//  exports
//    DllRegisterServer;
var
  ClassID, ServerKeyName,CategoryId :string;
Begin
  // �������� ����� ������
  ClassID   := GUIDToString(Class_Of_TSquareExtraction);
  // �������� ����� ���������
  CategoryId:= GUIDToString(CatParameterizedOperations);
  // ��������� ����� �������
  ServerKeyName := 'CLSID\' + ClassID + '\Implemented Categories';
  // �������� ���������� ����������� �����������
  Result:=ComServ.DllRegisterServer;
  // ������������ ���� ���������
  CreateRegKey(ServerKeyName ,' ','Categories');
  // ��������� �����������
  CreateRegKey(ServerKeyName+'\'+CategoryId,' ','AlfaLab �ategory of parameterized operations');
  // ��������� ��������� ��������� �����������
End;

*)

{ TSquareExtraction }

destructor TSquareExtraction.Destroy;
begin

  inherited;
end;

procedure TSquareExtraction.Initialize;
begin
  inherited;
  poClearParameters;
  poDataSquare.pNameParameter:='�����';
  poDataSquare.pHintParameter:='������� �����, �� �������� ����� �������� ���������� ������';
  poDataSquare.pRequired:=true;
  poDataSquare.pEnumeration:=333.33;
  poDataSquare.pEnumeration:=VarAsType(poDataSquare.pEnumeration, varDouble);
end;

function TSquareExtraction.pCheckParameters(Index: longint; pVerify: AnsiString;
  const Formatting: TFormatSettings): boolean;
var
 pVar : Variant;
begin
Result:=false;
try
  if (Length(pVerify)=0) or (Index<>1) then exit;
  pVar:=StrToFloat(pVerify, Formatting);
  Result:=true;
except
end;
end;

function TSquareExtraction.pGetEnumeration(Index: longint): Variant;
begin
  if Index<>1 then exit;
  Result:=poDataSquare.pEnumeration;
end;

function TSquareExtraction.pGetHintParameter(Index: longint): AnsiString;
begin
  if Index<>1 then exit;
  Result:=poDataSquare.pHintParameter;
end;

function TSquareExtraction.pGetNameParameter(Index: longint): AnsiString;
begin
  if Index<>1 then exit;
  Result:=poDataSquare.pNameParameter;
end;

function TSquareExtraction.pGetParameters(Index: longint): Variant;
begin
  if Index<>1 then exit;
  Result:=poDataSquare.pParameters;
end;

function TSquareExtraction.pGetRequired(Index: longint): boolean;
begin
  Result:=false;
  if Index<>1 then exit;
  Result:=poDataSquare.pRequired;
end;

function TSquareExtraction.poCheck: longint;
var
 chDataSquare   : Variant;
 FormatSettings : TFormatSettings;
begin
try
 try
  chDataSquare:=poDataSquare.pParameters;
  // ��������� ����������� ������
  FormatSettings.DecimalSeparator := '.';

  poDataSquare.pParameters:=StrToFloat('25.0', FormatSettings);

  Result:=poExecute;
  if Result<>0 then exit;
  // �������� ������������
  if (poDataSquare.pParameters<>5) then Result:=2;

 except
  Result:=-1;
 end;
finally
  poDataSquare.pParameters:=chDataSquare;
end;
end;

procedure TSquareExtraction.poClearParameters;
begin
 VarClear(poDataSquare.pParameters);
end;

function TSquareExtraction.poCountOfParameters: longint;
begin
 Result:=1;
end;

function TSquareExtraction.poExecute: longint;
begin
 Result:=0;
 try
   poDataSquare.pParameters:=Sqrt(poDataSquare.pParameters);
 except
  Result:=-1;
 end;
end;

function TSquareExtraction.poGetNameParameters: AnsiString;
begin
 Result:='���������� ����������� �����';
end;

function TSquareExtraction.poGetResult: Variant;
begin
 Result:=poDataSquare.pParameters;
end;

function TSquareExtraction.pSetParameters(Index: longint; pParam: AnsiString;
  const Formatting: TFormatSettings): boolean;
begin
Result:=false;
 if Index<>1 then exit;
try
 poDataSquare.pParameters:=StrToFloat(pParam , Formatting);
 Result:=true;
except
end;
end;

Initialization
  TComObjectFactory.Create (
            ComServer,
            TSquareExtraction,
            Class_Of_TSquareExtraction,
            'IParameterizedOperations',
            '���������� ����������� �����',
            ciMultiInstance,
            tmApartment);
Finalization
 Begin
 End;

end.
