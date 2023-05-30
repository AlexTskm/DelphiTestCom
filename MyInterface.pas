{$ALIGN OFF}
{$H+}
unit MyInterface;

interface

uses System.SysUtils, Activex;

const
// путь к приложению в реестре
 ThePathToTheApplication = 'SOFTWARE\AlfaLab\Test\';

// Объект увеличение даты
Class_Of_TIncTheDate:TCLSID=
'{12345678-0001-AADD-0001-080520232013}';

// Объект извлечение квадратного корня
Class_Of_TSquareExtraction:TCLSID=
'{12345678-0001-AADD-0001-110520230844}';


type
// Типы данных

// Данные параметризованной операции
PRecordingParameters =^TRecordingParameters;
TRecordingParameters = record
// Название параметра
  pNameParameter : AnsiString;
// Расширенное назначение параметра
  pHintParameter : AnsiString;
// Тип параметра(обязательный/необязательный)
  pRequired      : boolean;
// Тип параметра определяется функцией VarType(целочисленное число, перебор значений из массива и т.д)
  pEnumeration   : Variant;
// Параметр
  pParameters    : Variant;
end;

// Интерфейс IParameterizedOperations
// Содержит параметры, функции параметризованной операции
IParameterizedOperations = Interface(IUnknown)
['{12345678-0001-AADD-0000-060520231408}']
// Выполняемые функции параметризованной операции
// Проверить функции параметризованной операции (Возвращает: <>0 ошибка)
  function  poCheck:longint;StdCall;
// Выполнить функцию параметризованной операции (Возвращает: <>0 ошибка)
  function  poExecute:longint;StdCall;
// Очистить параметры функции параметризованной операции
  procedure  poClearParameters;StdCall;
// Функция выдаёт результат работы poExecute
  function poGetResult: Variant;StdCall;

// Выдаёт название параметризованной операции
  function poGetNameParameters:AnsiString;Stdcall;
// Количество параметров функции параметризованной операции (Возвращает: < 0 ошибка)
  function poCountOfParameters:longint;StdCall;
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
// Проверяет подходит ли строка pVerify для параметра Index начинается с 1 . В случае ошибки возвращает правильное pVerify
  function pCheckParameters(Index:longint; pVerify : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;
end;


implementation

end.
