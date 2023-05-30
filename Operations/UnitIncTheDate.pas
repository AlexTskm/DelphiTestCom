unit UnitIncTheDate;

interface

uses System.Classes, System.SysUtils, Comserv, ComObj, Activex, MyInterface;

type


  TIncTheDate = class(TComObject, IParameterizedOperations)
  private
  { Private declarations }
   // Дата
   poData : TRecordingParameters;
   // Увеличивается на сколько
   poIncrease : TRecordingParameters;
   // Единица увеличения
   poUnit : TRecordingParameters;

  protected
  { Protected declarations }

  public

//   Инициализация и деинициализация объектов
    procedure Initialize;override;
    destructor Destroy;override;
//
//  IParameterizedOperations
//
  // Выполняемые функции параметризованной операции
  // Проверить функции параметризованной операции (Возвращает: <>0 ошибка)
    function  poCheck:longint;virtual;StdCall;
  // Выполнить функцию параметризованной операции (Возвращает: <>0 ошибка)
    function  poExecute:longint;virtual;StdCall;
  // Очистить параметры функции параметризованной операции
    procedure  poClearParameters;virtual;StdCall;

  // Выдаёт название параметризованной операции
    function poGetNameParameters:AnsiString;virtual;Stdcall;
  // Параметры функции параметризованной операции (Возвращает: < 0 ошибка)
    function poCountOfParameters:longint;virtual;StdCall;
// Функция выдаёт результат работы poExecute
    function poGetResult: Variant;StdCall;
///////////////////////////////////////////////////////////////////////////////////////////////////
///  IRecordingParameters
  // Название параметра Index начинается с 1
    function pGetNameParameter(Index:longint) : AnsiString;StdCall;
  // Расширенное назначение параметра Index начинается с 1
    function pGetHintParameter(Index:longint) : AnsiString;StdCall;
  // Тип параметра(обязательный/необязательный) Index начинается с 1
    function pGetRequired(Index:longint) : boolean;StdCall;
  // Тип параметра определяется функцией VarType(целочисленное число, перебор значений из массива и т.д) Index начинается с 1
    function pGetEnumeration(Index:longint) : Variant;StdCall;
  // Получить Параметр Index начинается с 1
    function pGetParameters(Index:longint)  : Variant;StdCall;
  // Записать параметр возвращает <> 0, если ошибка Index начинается с 1
    function pSetParameters(Index:longint; pParam : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;
  // Проверяет Index начинается с 1
    function pCheckParameters(Index:longint; pVerify : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;

  end;

//////////////////////////////////////////////////////////////
// Регистрация компонента
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
// Регистрация компонента в категории компонентов
//////////////////////////////////////////////////////////////
// Регистрация компонента
//  Function DllRegisterServer:HResult; stdcall;
// Для экспорта
//  exports
//    DllRegisterServer;
var
  ClassID, ServerKeyName,CategoryId :string;
Begin
  // Пересчет имени класса
  ClassID   := GUIDToString(Class_Of_TIncTheDate);
  // Пересчет имени категории
  CategoryId:= GUIDToString(CatParameterizedOperations);
  // Получение имени сервера
  ServerKeyName := 'CLSID\' + ClassID + '\Implemented Categories';
  // Вызываем предыдущую регистрацию компонентов
  Result:=ComServ.DllRegisterServer;
  // Регистрируем свою категорию
  CreateRegKey(ServerKeyName ,' ','Categories');
  // Категория компонентов
  CreateRegKey(ServerKeyName+'\'+CategoryId,' ','AlfaLab Сategory of parameterized operations');
  // Получение менеджера категории компонентов
End;

*)

procedure TIncTheDate.Initialize;
begin
  inherited;
  poClearParameters;
  poData.pNameParameter:='Дата';
  poData.pHintParameter:='Введите дату, которую необходимо увеличить';
  poData.pRequired:=true;
  poData.pEnumeration:=Date;
  poData.pEnumeration:=VarAsType(poData.pEnumeration, varDate);

  poIncrease.pNameParameter:='Увеличить на';
  poIncrease.pHintParameter:='Введите положительное число, на которое необходимо увеличить';
  poIncrease.pRequired:=true;
  poIncrease.pEnumeration:=0;
  poIncrease.pEnumeration:=VarAsType(poIncrease.pEnumeration, varByte);

  poUnit.pNameParameter:='Единица';
  poUnit.pHintParameter:='Введите единицу, которую необходимо увеличить';
  poUnit.pRequired:=true;
  poUnit.pEnumeration:=VarArrayOf(['день','месяц','год']);

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
  // Установим проверочную дату
  FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  FormatSettings.DateSeparator := '-';
  FormatSettings.LongTimeFormat := 'hh:nn:ss';
  FormatSettings.TimeSeparator := ':';

  poData.pParameters:=StrToDateTime('2021-02-28 01:02:03.004', FormatSettings);
  poIncrease.pParameters:=1;
  poUnit.pParameters:=poUnit.pEnumeration[0];

  Result:=poExecute;
  if Result<>0 then exit;
  // Проверим правильность
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
// день
   poData.pParameters:= IncDay(poData.pParameters, poIncrease.pParameters);
 end else
 if poUnit.pParameters = poUnit.pEnumeration[1] then
 begin
// месяц
  poData.pParameters:= IncMonth(poData.pParameters, poIncrease.pParameters);
 end else
 if poUnit.pParameters = poUnit.pEnumeration[2] then
 begin
// год
  poData.pParameters:= IncYear(poData.pParameters, poIncrease.pParameters);
 end else
  Result:=1;
end;

function TIncTheDate.poGetNameParameters: AnsiString;
begin
 Result:='Увеличение даты';
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
            'Увеличение даты',
            ciMultiInstance,
            tmApartment);
Finalization
 Begin
 End;

end.

