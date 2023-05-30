{$ALIGN OFF}
{$H+}
unit CommonProcedure;

interface

uses System.Classes, Activex, Winapi.Windows;

// ������ ������� �������.
// ������� ������:
// ���� � ������� � ���������� � HKEY_LOCAL_MACHINE
// ������������ �������� ������ ��������� COM ��������
// true - ��� ����������
// false - ������ ������ � ��������

function GetFromRegistr (RegPath : string; var Elements :  TInterfaceList): boolean;

// �������� ������ � ������
function SaveFromRegistr (RegPath : string; Class_Of_Object:TCLSID): boolean;

// ����������� COM �������
function RegFilePr(hWnd: HWND; FName: String): Boolean;


implementation

uses ShellApi, System.SysUtils, Dialogs, ComObj, Registry, MyInterface;

// ����� � �������, ���� <>0 �� ������ ������ ������������
const KeyUse = 'USE';
const
  RegSvr32Str='RegSvr32.exe';


// ������ ������� �������.
// ������������ �������� ������ ��������� COM ��������
// true - ��� ����������
// false - ������ ������ � ��������

function GetFromRegistr (RegPath : string; var Elements :  TInterfaceList): boolean;
var
  Reg : TRegistry; // ������
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
// ������� ������ ��� �������� �� �������
        OwnerSection := TStringList.Create;
        if not Assigned (OwnerSection) then Exit;
        Reg:=TRegistry.Create(KEY_READ);
        if not Assigned (Reg) then Exit;
        Reg.RootKey:=HKEY_LOCAL_MACHINE;
// ���� ��� ������ ���� ������������ �� � ���
        If not Reg.KeyExists(RegPath) Then Exit;
// ��������� ����� ����������
        if not Reg.OpenKey(RegPath,False) Then Exit;
// ������ ������ ����������������� ��������
        Reg.GetKeyNames(OwnerSection);
// ���������� ����������������� ��������
        if OwnerSection.Count = 0 then Exit;
        for N := 0 to OwnerSection.Count-1 do
        begin
          Reg.CloseKey;
// ��������, ��� �� ������-������ ��� ��������� ���� Use
// ���� ���� Use = 0, �� ������ ����������������� �������� �� ���������
          Reg.OpenKey (RegPath + '\'+ OwnerSection.Strings[N], FALSE);
          if Reg.ValueExists (KeyUse) then
            if Reg.ReadInteger (KeyUse) = 0 then
              continue;
          // ������ COM ������ � ��������� ��� � ������ ����������� ����������������� ��������
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
  Reg : TRegistry; // ������
begin
      Result := FALSE;
      Reg:=nil;
      try
        try
  // ������� ������ ��� �������� �� �������
          Reg:=TRegistry.Create(KEY_READ);
          if not Assigned (Reg) then Exit;
          Reg.RootKey:=HKEY_LOCAL_MACHINE;
          RegPath:=RegPath + GUIDToString(Class_Of_Object);
  // ���������/������  ����� ����������
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
