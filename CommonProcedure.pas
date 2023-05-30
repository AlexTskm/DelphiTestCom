{$ALIGN OFF}
{$H+}
unit CommonProcedure;

interface

uses System.Classes, Activex, Winapi.Windows;

// Чтение записей реестра.
// Входные данные:
// Путь в реестре к приложению в HKEY_LOCAL_MACHINE
// Возвращаемые значения список созданных COM объектов
// true - все получилось
// false - ошибка работы с реестром

function GetFromRegistr (RegPath : string; var Elements :  TInterfaceList): boolean;

// Записать запись в реестр
function SaveFromRegistr (RegPath : string; Class_Of_Object:TCLSID): boolean;

// Регистрация COM объекта
function RegFilePr(hWnd: HWND; FName: String): Boolean;


implementation

uses ShellApi, System.SysUtils, Dialogs, ComObj, Registry, MyInterface;

// Ключь в реестре, если <>0 то данный объект используется
const KeyUse = 'USE';
const
  RegSvr32Str='RegSvr32.exe';


// Чтение записей реестра.
// Возвращаемые значения список созданных COM объектов
// true - все получилось
// false - ошибка работы с реестром

function GetFromRegistr (RegPath : string; var Elements :  TInterfaceList): boolean;
var
  Reg : TRegistry; // Реестр
  N   : integer;
  OwnerSection : TStringList;
  IE  : IUnknown;
  iEp : IParameterizedOperations;
begin
      Result := FALSE;
      Reg:=nil;
      OwnerSection := nil;
      Elements := nil;
      try
// Создаем объект для ползания по реестру
        OwnerSection := TStringList.Create;
        if not Assigned (OwnerSection) then Exit;
        Reg:=TRegistry.Create(KEY_READ);
        if not Assigned (Reg) then Exit;
        Reg.RootKey:=HKEY_LOCAL_MACHINE;
// Если нет такого пути возвращаемся ни с чем
        If not Reg.KeyExists(RegPath) Then Exit;
// Открываем ключи приложения
        if not Reg.OpenKey(RegPath,False) Then Exit;
// Читаем список параметризованных операций
        Reg.GetKeyNames(OwnerSection);
// Количество параметризованных операций
        if OwnerSection.Count = 0 then Exit;
        for N := 0 to OwnerSection.Count-1 do
        begin
          Reg.CloseKey;
// Проверим, нет ли какого-нибудь там параметра типа Use
// Если есть Use = 0, то данную параметризованную операцию не загружаем
          Reg.OpenKey (RegPath + '\'+ OwnerSection.Strings[N], FALSE);
          if Reg.ValueExists (KeyUse) then
            if Reg.ReadInteger (KeyUse) = 0 then
              continue;
          // Создаём COM объект и добавляем его в список загруженных параметризованных операций
          if not Assigned(Elements) then Elements:=TInterfaceList.Create;
          IE:=CreateComObject(StringtoGuid(OwnerSection.Strings[N]));
          if Assigned(IE) then
          begin
           try
            iEp := IE as IParameterizedOperations;
           except
            iEp :=nil;
           end;
           if Assigned(IEp) then
             ELements.Add(IE);
          end;
        end;
        Result := TRUE;
      finally
        if Assigned (Reg) then
        begin
          Reg.CloseKey;
          Reg.Free;
        end;
        if Assigned (OwnerSection) then OwnerSection.Free;
      end;
end;


function SaveFromRegistr (RegPath : string; Class_Of_Object:TCLSID): boolean;
var
  Reg : TRegistry; // Реестр
begin
      Result := FALSE;
      Reg:=nil;
      try
        try
  // Создаем объект для ползания по реестру
          Reg:=TRegistry.Create(KEY_READ);
          if not Assigned (Reg) then Exit;
          Reg.RootKey:=HKEY_LOCAL_MACHINE;
          RegPath:=RegPath + GUIDToString(Class_Of_Object);
  // Открываем/создаём  ключи приложения
          if not Reg.OpenKey(RegPath,true) Then Exit;
//            if not Reg.ValueExists('Use') then
//              Reg.WriteInteger('Use',1);
          Result := TRUE;
        except
        end;
      finally
        if Assigned (Reg) then
        begin
          Reg.CloseKey;
          Reg.Free;
        end;
      end;
end;

function RegFilePr(hWnd: HWND; FName: String): Boolean;
begin
  Result:=False;
  try
    if ShellExecute(hWnd,
                 nil,
                 RegSvr32Str,
                 PChar(FName),
                 nil,
                 SW_SHOWNORMAL)<=32 then Exit;
    Result:=True;
  except

  end;
end;


end.
