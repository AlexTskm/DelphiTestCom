unit UnitSquare;

interface

uses System.Classes, System.SysUtils, Comserv, ComObj, Activex, MyInterface;


type


  TSquareExtraction = class(TComObject, IParameterizedOperations)
  private
  { Private declarations }
   // Число из которого будет извлекаться корень
   poDataSquare : TRecordingParameters;

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
  ClassID   := GUIDToString(Class_Of_TSquareExtraction);
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

{ TSquareExtraction }

destructor TSquareExtraction.Destroy;
begin

  inherited;
end;

procedure TSquareExtraction.Initialize;
begin
  inherited;
  poClearParameters;
  poDataSquare.pNameParameter:='Число';
  poDataSquare.pHintParameter:='Введите число, из которого будет извлечён квадратный корень';
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
  // Установим проверочные данные
  FormatSettings.DecimalSeparator := '.';

  poDataSquare.pParameters:=StrToFloat('25.0', FormatSettings);

  Result:=poExecute;
  if Result<>0 then exit;
  // Проверим правильность
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
 Result:='Извлечение квадратного корня';
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
            'Извлечение квадратного корня',
            ciMultiInstance,
            tmApartment);
Finalization
 Begin
 End;

end.
