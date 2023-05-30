unit UnitIncTheDate;

interface

uses System.Classes, System.SysUtils, Comserv, ComObj, Activex, MyInterface;

type


  TIncTheDate = class(TComObject, IParameterizedOperations)
  private
  { Private declarations }
   // ����
   poData : TRecordingParameters;
   // ������������� �� �������
   poIncrease : TRecordingParameters;
   // ������� ����������
   poUnit : TRecordingParameters;

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
  ClassID   := GUIDToString(Class_Of_TIncTheDate);
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

procedure TIncTheDate.Initialize;
begin
  inherited;
  poClearParameters;
  poData.pNameParameter:='����';
  poData.pHintParameter:='������� ����, ������� ���������� ���������';
  poData.pRequired:=true;
  poData.pEnumeration:=Date;
  poData.pEnumeration:=VarAsType(poData.pEnumeration, varDate);

  poIncrease.pNameParameter:='��������� ��';
  poIncrease.pHintParameter:='������� ������������� �����, �� ������� ���������� ���������';
  poIncrease.pRequired:=true;
  poIncrease.pEnumeration:=0;
  poIncrease.pEnumeration:=VarAsType(poIncrease.pEnumeration, varByte);

  poUnit.pNameParameter:='�������';
  poUnit.pHintParameter:='������� �������, ������� ���������� ���������';
  poUnit.pRequired:=true;
  poUnit.pEnumeration:=VarArrayOf(['����','�����','���']);

end;

destructor TIncTheDate.Destroy;
begin
try

finally
  inherited
end;
end;


function TIncTheDate.pCheckParameters(Index: longint; pVerify : AnsiString; const Formatting : TFormatSettings): boolean;
var
 pVar : Variant;
begin
Result:=false;
try
if Length(pVerify)=0 then exit;
 case Index of
 1:begin
     pVar:=StrToDate(pVerify, Formatting);
     Result:=true;
   end;
 2:begin
    pVar:=StrToInt(pVerify);
    if (pVar<0) or(pVar>256) then exit;
    Result:=true;
   end;
 3:begin
    if (poUnit.pEnumeration[0]=pVerify) or (poUnit.pEnumeration[1]=pVerify) or (poUnit.pEnumeration[2]=pVerify)then
      Result:=true;
   end;
 end;
except
end;
end;

function TIncTheDate.pGetEnumeration(Index: longint): Variant;
begin
 case Index of
 1: Result:=poData.pEnumeration;
 2: Result:=poIncrease.pEnumeration;
 3: Result:=poUnit.pEnumeration;
 end;
end;

function TIncTheDate.pGetHintParameter(Index: longint): AnsiString;
begin
 Result:='';
 case Index of
 1: Result:=poData.pHintParameter;
 2: Result:=poIncrease.pHintParameter;
 3: Result:=poUnit.pHintParameter;
 end;
end;

function TIncTheDate.pGetNameParameter(Index: longint): AnsiString;
begin
 Result:='';
 case Index of
 1: Result:=poData.pNameParameter;
 2: Result:=poIncrease.pNameParameter;
 3: Result:=poUnit.pNameParameter;
 end;
end;

function TIncTheDate.pGetParameters(Index: longint): Variant;
begin
 case Index of
 1: Result:=poData.pParameters;
 2: Result:=poIncrease.pParameters;
 3: Result:=poUnit.pParameters;
 end;
end;

function TIncTheDate.pGetRequired(Index: longint): boolean;
begin
 Result:=false;
 case Index of
 1: Result:=poData.pRequired;
 2: Result:=poIncrease.pRequired;
 3: Result:=poUnit.pRequired;
 end;
end;

function TIncTheDate.poCheck: longint;
var
 chData, chIncrease, chUnit : Variant;
 FormatSettings : TFormatSettings;
begin
try
 try
  chData:=poData.pParameters;
  chIncrease:=poIncrease.pParameters;
  chUnit:=poUnit.pParameters;
  // ��������� ����������� ����
  FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  FormatSettings.DateSeparator := '-';
  FormatSettings.LongTimeFormat := 'hh:nn:ss';
  FormatSettings.TimeSeparator := ':';

  poData.pParameters:=StrToDateTime('2021-02-28 01:02:03.004', FormatSettings);
  poIncrease.pParameters:=1;
  poUnit.pParameters:=poUnit.pEnumeration[0];

  Result:=poExecute;
  if Result<>0 then exit;
  // �������� ������������
  if (DayOf(poData.pParameters)<>1) or
     (MonthOf(poData.pParameters)<>3) or
     (YearOf(poData.pParameters)<>2021) then Result:=2;

 except
  Result:=-1;
 end;
finally
  poData.pParameters:=chData;
  poIncrease.pParameters:=chIncrease;
  poUnit.pParameters:=chUnit;
end;
end;

procedure TIncTheDate.poClearParameters;
begin
 VarClear(poData.pParameters);
 VarClear(poIncrease.pParameters);
 VarClear(poUnit.pParameters);
end;

function TIncTheDate.poCountOfParameters: longint;
begin
 Result:=3;
end;

function TIncTheDate.poExecute: longint;
begin
 Result:=0;
 if poUnit.pParameters = poUnit.pEnumeration[0] then
 begin
// ����
   poData.pParameters:= IncDay(poData.pParameters, poIncrease.pParameters);
 end else
 if poUnit.pParameters = poUnit.pEnumeration[1] then
 begin
// �����
  poData.pParameters:= IncMonth(poData.pParameters, poIncrease.pParameters);
 end else
 if poUnit.pParameters = poUnit.pEnumeration[2] then
 begin
// ���
  poData.pParameters:= IncYear(poData.pParameters, poIncrease.pParameters);
 end else
  Result:=1;
end;

function TIncTheDate.poGetNameParameters: AnsiString;
begin
 Result:='���������� ����';
end;

function TIncTheDate.poGetResult: Variant;
begin
 Result:=poData.pParameters;
end;

function TIncTheDate.pSetParameters(Index: longint; pParam : AnsiString; const Formatting : TFormatSettings) : boolean;
begin
Result:=false;
if Length(pParam)=0 then exit;
try
   case Index of
   1:
     begin
       poData.pParameters:=StrToDate(pParam, Formatting);
       Result:=true;
     end;
   2:
     begin
       poIncrease.pParameters:=StrToInt(pParam);
       Result:=true;
     end;
   3:
     if (poUnit.pEnumeration[0]=pParam) or (poUnit.pEnumeration[1]=pParam) or (poUnit.pEnumeration[2]=pParam)then
     begin
       poUnit.pParameters:=pParam;
       Result:=true;
     end;
   end;
except
end;
end;

Initialization
  TComObjectFactory.Create (
            ComServer,
            TIncTheDate,
            Class_Of_TIncTheDate,
            'IParameterizedOperations',
            '���������� ����',
            ciMultiInstance,
            tmApartment);
Finalization
 Begin
 End;

end.

