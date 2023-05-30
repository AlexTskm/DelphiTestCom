unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, MyInterface;

type
  TFMain = class(TForm)
    Panel1: TPanel;
    btnCheck: TButton;
    btnExecute: TButton;
    btnClearParameters: TButton;
    btnExit: TButton;
    Panel2: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    outResult: TLabel;
    GroupBox1: TGroupBox;
    cbAction: TComboBox;
    GroupBox2: TGroupBox;
    procedure FormCreate(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCheckClick(Sender: TObject);
    procedure cbActionChange(Sender: TObject);
    procedure btnExecuteClick(Sender: TObject);
    procedure EditEnter(Sender: TObject);
    procedure btnClearParametersClick(Sender: TObject);
  private
    iObjectParameters : integer;   // Активный объект параметризованных операций
    pComponentList : Array of Array of variant; // Список компонентов
    Elements :  TInterfaceList;    // Список объектов параметризованных операций
    formatSettings : TFormatSettings; // форматирование данных

    procedure Error(sError: string);       // Выводит сообщение об ошибке
    procedure pMessage(sMess: string);     // Выводит сообщение
    procedure CheckActivButton;            // Делает активными или не активными кнопки действия
    function ResultToString(pVar : Variant): string; // Преобразует переменную Variant в string
    function CreateComponentParamOperations(id : integer): longint; // Создание и размещение компонентов для параметризованной операции
    function CreateCompParam(var pLeft , pTop:integer; idx, index : integer; iEp : IParameterizedOperations; var Comp : TComponent): longint; // Создание компонента
    function GetTheCorrectString(idx: integer; iEp: IParameterizedOperations): String; // Выдаёт начальную установку для параметра
    function SaveParamOperations(idx: integer):boolean; // Записать введённые параметры в объект
    function VisibleCompParam(idx: integer; fV : boolean):boolean; // Включить/ выключить видимость компонентов параметрической операции
    function ClearParamOperations(idx: integer): boolean; // Очистить параметры
    { Private declarations }
  public
    { Public declarations }
  end;

  TMyTEdit = class(TEdit)
    private
     index: integer;
     iEParameterized : IParameterizedOperations;
     formatSettings : TFormatSettings; // форматирование данных
    public
     constructor Create(AOwner: TComponent); override;
     procedure Init(idx: integer; iEp: IParameterizedOperations; fS : TFormatSettings);
     function GetIndex:integer;
     function GetiEParameterized : IParameterizedOperations;
     function SetParamOperations : boolean;
  end;

  TMyTComboBox = class(TComboBox)
    private
     index: integer;
     iEParameterized : IParameterizedOperations;
     formatSettings : TFormatSettings; // форматирование данных
    public
     constructor Create(AOwner: TComponent); override;
     procedure Init(idx: integer; iEp: IParameterizedOperations; fS : TFormatSettings);
     function GetIndex:integer;
     function GetiEParameterized : IParameterizedOperations;
     function SetParamOperations : boolean;
  end;

var
  FMain: TFMain;

implementation

uses CommonProcedure;

{$R *.dfm}


procedure TFMain.btnCheckClick(Sender: TObject);
var
 iEp : IParameterizedOperations;
begin
 if cbAction.ItemIndex<0 then exit;
 if cbAction.ItemIndex>ELements.Count-1 then exit;
 try
  iEp := ELements[cbAction.ItemIndex] as IParameterizedOperations;
 except
  Error('Ошибка объекта параметризованных операций. Не поддерживает IParameterizedOperations');
  iEp := nil;
  exit;
 end;
 if Assigned(iEp) then
 begin
  if IEp.poCheck<>0 then
    Error('Ошибка проверки функции параметризованной операции.')
  else
   pMessage('Успешно выполнена проверка функции параметризованной операции ');
 end;
end;

procedure TFMain.btnClearParametersClick(Sender: TObject);
var
 iEp : IParameterizedOperations;
begin
 if cbAction.ItemIndex<0 then exit;
 if cbAction.ItemIndex>ELements.Count-1 then exit;
 try
  iEp := ELements[cbAction.ItemIndex] as IParameterizedOperations;
 except
  Error('Ошибка объекта параметризованных операций. Не поддерживает IParameterizedOperations');
  iEp := nil;
  exit;
 end;
 if Assigned(iEp) then
 begin
  iEp.poClearParameters;
  if not ClearParamOperations(cbAction.ItemIndex) then
    Error('Ошибка очистки параметров параметризованной операций');
 end;
end;

procedure TFMain.btnExecuteClick(Sender: TObject);
var
 iEp : IParameterizedOperations;
begin
 if cbAction.ItemIndex<0 then exit;
 if cbAction.ItemIndex>ELements.Count-1 then exit;
 try
  iEp := ELements[cbAction.ItemIndex] as IParameterizedOperations;
 except
  Error('Ошибка объекта параметризованных операций. Не поддерживает IParameterizedOperations');
  iEp := nil;
  exit;
 end;
 if Assigned(iEp) then
 begin
// Снача запишем параметры
  if not SaveParamOperations(cbAction.ItemIndex) then
    Error('Ошибка записи параметров параметризованной операции.')
  else
  if IEp.poExecute<>0 then
    Error('Ошибка выполнения функции параметризованной операции.')
  else
  begin
    pMessage(ResultToString(IEp.poGetResult));
  end;
 end;
end;

procedure TFMain.btnExitClick(Sender: TObject);
begin
  self.Close;
end;

procedure TFMain.cbActionChange(Sender: TObject);
var
 i:integer;
begin
// Установим активным объект параметризованных операций
// выведем его параметры
  CheckActivButton;
// Делаем невидимыми все параметры(лучше, конечно делать невидимым только предыдущий, для этого примера это не принципиально)
  for i := 0 to cbAction.Items.Count-1 do
    VisibleCompParam(i, false);
// Создание компонентов для параметров параметризованных операций
  CreateComponentParamOperations(cbAction.ItemIndex);
end;

procedure TFMain.CheckActivButton;
begin
 if cbAction.ItemIndex>=0 then
 begin
  btnCheck.Enabled:=true;
  btnExecute.Enabled:=true;
  btnClearParameters.Enabled:=true;
 end else
 begin
  btnCheck.Enabled:=false;
  btnExecute.Enabled:=false;
  btnClearParameters.Enabled:=false;
 end;
end;

function TFMain.CreateComponentParamOperations(id: integer): longint;
const
 LeftMargin = 20;
 TopMargin = 20;
var
 i, j, vType, td, th, pLeft, pTop, hAlign:integer;
 iEp : IParameterizedOperations;
 Comp : TComponent;
 pLabel : TLabel;
 s:string;
begin
  Result:=1;
  if (id<0) or (id>=Elements.Count) then exit;
  try
   iEp := ELements[id] as IParameterizedOperations;
  except
   Error('Ошибка объекта параметризованных операций. Не поддерживает IParameterizedOperations');
   iEp := nil;
   Result:=-1;
   exit;
  end;
  if Assigned(iEp) then
  begin
    if Length(pComponentList[id])<>(IEp.poCountOfParameters*2) then
      SetLength(pComponentList[id] , IEp.poCountOfParameters*2);
    pLeft:=LeftMargin;
    pTop:=TopMargin;
    j:=-1;
    for I := 1 to IEp.poCountOfParameters do
    begin
     inc(j);
     vType:=VarType(pComponentList[id, j]);
     if vType<>varString then
     begin
      // Создаём и размещаем компонент
      pLabel:=TLabel.Create(GroupBox2);
      pLabel.Parent:=GroupBox2;
      pLabel.Caption:=IEp.pGetNameParameter(i) + ':';
      pLabel.Name:='pLabel' + IntToStr(id) + IntToStr(i);
      pComponentList[id, j]:=pLabel.Name;
      pComponentList[id, j]:=VarAsType(pComponentList[id, j], varString);
      td:=pLabel.Width;
      th:=pLabel.Height;
      pLabel.Left:=pLeft;
      pLabel.Top:=pTop;
      pLabel.Visible:=true;
      pLeft:=pLeft + td;
      hAlign:=(pTop+(th div 2));
      if CreateCompParam(pLeft, hAlign , id, i, IEp , Comp)<> 0 then
      begin
       Error('Ошибка создания компонента параметра у объекта параметризованных операций.');
       exit;
      end;
      pTop:=hAlign-(th div 2);
      s:=Comp.Name;
      inc(j);
      pComponentList[id, j]:=S;
      pComponentList[id, j]:=VarAsType(pComponentList[id, j], varString);
      pLeft:=pLeft+LeftMargin;
//      pTop:=pTop+th;
     end else
     begin
      // Нужно просто включить компонентам видимость
      if VisibleCompParam(id,true) then
      begin
       Result:=0;
       exit;
      end;
     end;
    end;
    Result:=0;
  end;
 end;

function TFMain.CreateCompParam(var pLeft , pTop:integer; idx, index : integer; iEp : IParameterizedOperations; var Comp: TComponent): longint;
var
 i, vType:integer;
 Edit: TMyTEdit;
 ComboBox: TMyTComboBox;
 s : string;
 pVar : Variant;

function GetEditTextWidth(Edit: TEdit; t : string):longint;
var
  Canvas: TControlCanvas;
begin
// Определение ширины компонента
Canvas := TControlCanvas.Create;
try
  Canvas.Control := Edit;
  Canvas.Font.Assign(Edit.Font);
  Result:=Canvas.TextWidth(t);
finally
  Canvas.Free;
end;
end;

function GetComboBoxTextWidth(ComboBox:TComboBox; t : string):longint;
var
  Canvas: TControlCanvas;
begin
// Определение ширины компонента
Canvas := TControlCanvas.Create;
try
  Canvas.Control := ComboBox;
  Canvas.Font.Assign(ComboBox.Font);
  Result:=Canvas.TextWidth(t);
finally
  Canvas.Free;
end;
end;


begin
Result:=0;
if not Assigned(iEp) then
begin
 Result:=-1;
 exit;
end;
pVar:=IEp.pGetEnumeration(index);
vType:=VarType(pVar);
case vType of
varDate:
begin
  Edit:=TMyTEdit.Create(GroupBox2);
  Edit.Parent:=GroupBox2;
  Edit.Name:='pParams' + IntToStr(idx) + IntToStr(index);
  Edit.Init(index, iEp, formatSettings);
  Edit.OnExit:=EditEnter;
  Edit.Left:=pLeft;
  Edit.Top:=pTop-(Edit.Height div 2);
  Edit.Text:=' ' + GetTheCorrectString(index, iEp) + ' ';
  Edit.Width:=GetEditTextWidth(Edit,' dd.mm.yyyy ');
  Edit.Hint:=IEp.pGetHintParameter(index)+ '. Формат записи(' + Edit.Text+')';
  Edit.ShowHint:=true;
  Edit.Visible:=true;
//  pTop:=pTop + Edit.Height;
  pLeft:=pLeft + Edit.Width;
  Comp:=Edit;
end;
varByte:
begin
  Edit:=TMyTEdit.Create(GroupBox2);
  Edit.Parent:=GroupBox2;
  Edit.Name:='pParams' + IntToStr(idx) + IntToStr(index);
  Edit.Init(index, iEp, formatSettings);
  Edit.OnExit:=EditEnter;
  Edit.Left:=pLeft;
  Edit.Top:=pTop-(Edit.Height div 2);
  Edit.Text:=GetTheCorrectString(index, iEp);
  Edit.Width:=GetEditTextWidth(Edit,' 000 ');
  Edit.Hint:=IEp.pGetHintParameter(index)+ '. Целочисленное число от 0 до 256';
  Edit.ShowHint:=true;
  Edit.Visible:=true;
//  pTop:=pTop + Edit.Height;
  pLeft:=pLeft + Edit.Width;
  Comp:=Edit;
end;
varSingle, varDouble, varCurrency:
begin
  Edit:=TMyTEdit.Create(GroupBox2);
  Edit.Parent:=GroupBox2;
  Edit.Name:='pParams' + IntToStr(idx) + IntToStr(index);
  Edit.Init(index, iEp, formatSettings);
  Edit.OnExit:=EditEnter;
  Edit.Left:=pLeft;
  Edit.Top:=pTop-(Edit.Height div 2);
  Edit.Text:=GetTheCorrectString(index, iEp);
  Edit.Width:=GetEditTextWidth(Edit,' OOOOOOOOOO ');
  Edit.Hint:=IEp.pGetHintParameter(index)+ '. Число с плавающей запятой';
  Edit.ShowHint:=true;
  Edit.Visible:=true;
//  pTop:=pTop + Edit.Height;
  pLeft:=pLeft + Edit.Width;
  Comp:=Edit;
end;
8204{varArray}:
begin
  ComboBox:=TMyTComboBox.Create(GroupBox2);
  ComboBox.Parent:=GroupBox2;
  ComboBox.Name:='pParams' + IntToStr(idx) + IntToStr(index);
  ComboBox.Init(index, iEp, formatSettings);
  ComboBox.Left:=pLeft;
  ComboBox.Top:=pTop-(ComboBox.Height div 2);
  s:='';
  for I := VarArrayLowBound(pVar,1) to VarArrayHighBound(pVar,1) do
  begin
    ComboBox.Items.Add(pVar[i]);
    if Length(s)<Length(pVar[i]) then
      s:=pVar[i];
  end;
  if Length(s)>0 then
    ComboBox.ItemIndex:=0;
  ComboBox.Width:=GetComboBoxTextWidth(ComboBox,s+'OOO');
  ComboBox.Hint:=IEp.pGetHintParameter(index);
  ComboBox.ShowHint:=true;
  ComboBox.Style:= csDropDownList;
  ComboBox.Visible:=true;
//  pTop:=pTop + ComboBox.Height;
  pLeft:=pLeft + ComboBox.Width;
  Comp:=ComboBox;
end;
end;
end;

procedure TFMain.EditEnter(Sender: TObject);
var
 IEp : IParameterizedOperations;
 idx : integer;
 S : AnsiString;
begin
// Процедура обработки Edit
with (Sender as TMyTEdit) do
begin
  iEp :=GetiEParameterized;
  idx:=GetIndex;
  if not Assigned(iEp) then exit;
  S:=Text;
  if not iEp.pCheckParameters(idx, S, formatSettings) then
  begin
    S:=GetTheCorrectString(idx,iEp);
    Text:=S;
    SetFocus;
    Error('Введите правильное значение');
  end;
end;
end;

procedure TFMain.Error(sError: string);
begin
 outResult.Font.Color:= clRed;
 outResult.Caption:=sError;
 self.Repaint;
end;

procedure TFMain.FormCreate(Sender: TObject);
var
 i, j:integer;
 iEp : IParameterizedOperations;
 s:string;
begin
// Устанавливаем минимальные размеры формы
self.Constraints.MinWidth:= Panel1.Left + btnExit.Left + btnExit.ClientWidth + (self.Width - self.ClientWidth);

GetLocaleFormatSettings(LOCALE_SYSTEM_DEFAULT, formatSettings);

iObjectParameters:=-1;
Elements:=TInterfaceList.Create;

if GetFromRegistr(ThePathToTheApplication,Elements) then
begin
 SetLength(pComponentList , ELements.Count);
 j:=-1;
 for i:= 0 to ELements.Count-1 do
 begin
  try
   iEp := ELements[i] as IParameterizedOperations;
  except
   Error('Ошибка объекта параметризованных операций. Не поддерживает IParameterizedOperations');
   iEp := nil;
  end;
  if Assigned(iEp) then
  begin
    s:=IEp.poGetNameParameters;
    cbAction.Items.Add(s);
    inc(j);
    SetLength(pComponentList[j] , IEp.poCountOfParameters * 2);
  end;
 end;
end
else
  Error('Ошибка загрузки объектов параметризованных операций');
end;

procedure TFMain.FormDestroy(Sender: TObject);
begin
 if Assigned(Elements) then Elements.Free;
end;

function TFMain.GetTheCorrectString(idx: integer;
  iEp: IParameterizedOperations): String;
var
 pVar : Variant;
 vType:integer;
begin
pVar:=IEp.pGetEnumeration(idx);
vType:=VarType(pVar);
case vType of
varDate:
begin
  DateTimeToString(Result, formatSettings.ShortDateFormat, pVar);
end;
varByte:
begin
  Result:=IntToStr(pVar);
end;
varSingle, varDouble, varCurrency:
begin
  Result:=FloatToStr(pVar, formatSettings);
end;
8204{varArray+varVariant}:
begin
  Result:=pVar[0];
end;
end;
end;

procedure TFMain.pMessage(sMess: string);
begin
 outResult.Font.Color:= clWindowText;
 outResult.Caption:=sMess;
 self.Repaint;
end;

function TFMain.ResultToString(pVar: Variant): string;
var
 vType : integer;
begin
Result:='';
vType:=VarType(pVar);
case vType of
varSingle, varDouble, varCurrency:	  // Значение с плавающей запятой.
begin
 Result:=FloatToStr(pVar, formatSettings);
end;
varDate:	    //  $0007	Значение даты и времени (тип TDateTime).
begin
  DateTimeToString(Result, formatSettings.ShortDateFormat, pVar);
end;
varByte:	    //  $0011	8-ми битовое беззнаковое целочислен-ное значение (тип Byte).
begin
  Result:=IntToStr(pVar);
end;
varString:	  //  $0100	Ссылка на динамически распределен-ную Pascal-строку (тип AnsiString).
begin
  Result:=pVar;
end;
else
 Result:='Данный тип не обрабатывается';
end;
end;

function TFMain.SaveParamOperations(idx: integer): boolean;
var
 i:integer;
 NameComp:string;
 Comp: TComponent;
 Edit: TMyTEdit;
 ComboBox:TMyTComboBox;
begin
Result:=false;
if Length(pComponentList[idx])=0 then exit;

for i := 0 to Length(pComponentList[idx])-1 do
begin
  NameComp:=pComponentList[idx,i];
  if Length(NameComp)=0 then continue;
  Comp:=GroupBox2.FindComponent(NameComp);
  if Comp<>nil then
  begin
    if Comp is TMyTEdit then
    begin
      Edit:=Comp as TMyTEdit;
      if Edit.SetParamOperations then Result:=true;
    end else
    if Comp is TMyTComboBox then
    begin
      ComboBox:=Comp as TMyTComboBox;
      if ComboBox.SetParamOperations then Result:=true;
    end;
  end;
end;
end;

function TFMain.ClearParamOperations(idx: integer): boolean;
var
 i:integer;
 NameComp:string;
 Comp: TComponent;
 Edit: TMyTEdit;
 ComboBox:TMyTComboBox;
begin
Result:=false;
if Length(pComponentList[idx])=0 then exit;

for i := 0 to Length(pComponentList[idx])-1 do
begin
  NameComp:=pComponentList[idx,i];
  if Length(NameComp)=0 then continue;
  Comp:=GroupBox2.FindComponent(NameComp);
  if Comp<>nil then
  begin
    if Comp is TMyTEdit then
    begin
      Edit:=Comp as TMyTEdit;
      Edit.Text:=GetTheCorrectString(Edit.index, Edit.iEParameterized);
      Result:=true;
    end else
    if Comp is TMyTComboBox then
    begin
      ComboBox:=Comp as TMyTComboBox;
      ComboBox.ItemIndex:=0;
      Result:=true;
    end;
  end;
end;
end;


function TFMain.VisibleCompParam(idx: integer; fV: boolean): boolean;
var
 i:integer;
 NameComp:string;
 Comp: TComponent;
 Edit: TMyTEdit;
 ComboBox:TMyTComboBox;
 Lab:TLabel;
begin
Result:=false;
if Length(pComponentList[idx])=0 then exit;

for i := 0 to Length(pComponentList[idx])-1 do
begin
  NameComp:=pComponentList[idx,i];
  if Length(NameComp)=0 then continue;
  Comp:=GroupBox2.FindComponent(NameComp);
  if Comp<>nil then
  begin
    if Comp is TMyTEdit then
    begin
      Edit:=Comp as TMyTEdit;
      Edit.Visible:=fV;
      Result:=true;
    end else
    if Comp is TMyTComboBox then
    begin
      ComboBox:=Comp as TMyTComboBox;
      ComboBox.Visible:=fV;
      Result:=true;
    end  else
    if Comp is TLabel then
    begin
      Lab:=Comp as TLabel;
      Lab.Visible:=fV;
      Result:=true;
    end;
  end;
end;
end;

{ TMyTEdit }

constructor TMyTEdit.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  iEParameterized:=nil;
  index:=-1;
end;

function TMyTEdit.GetiEParameterized: IParameterizedOperations;
begin
 Result:=iEParameterized;
end;

function TMyTEdit.GetIndex: integer;
begin
 Result:=index;
end;

procedure TMyTEdit.Init(idx: integer; iEp: IParameterizedOperations; fS : TFormatSettings);
begin
 index:=idx;
 iEParameterized:=iEp;
 formatSettings:=fS;
end;

function TMyTEdit.SetParamOperations: boolean;
var
 S : String;
begin
 Result:=false;
 if not Assigned(iEParameterized) or (index<0) then exit;
 S:=self.Text;
 if not iEParameterized.pCheckParameters(index, S, formatSettings) then exit;
 if not iEParameterized.pSetParameters(index, S, formatSettings) then exit;
 Result:=true;
end;

{ TMyTComboBox }

constructor TMyTComboBox.Create(AOwner: TComponent);
begin
  inherited;
  inherited Create(AOwner);
  iEParameterized:=nil;
  index:=-1;
end;

function TMyTComboBox.GetiEParameterized: IParameterizedOperations;
begin
 Result:=iEParameterized;
end;

function TMyTComboBox.GetIndex: integer;
begin
 Result:=index;
end;

procedure TMyTComboBox.Init(idx: integer; iEp: IParameterizedOperations; fS : TFormatSettings);
begin
 index:=idx;
 iEParameterized:=iEp;
 formatSettings:=fS;
end;

function TMyTComboBox.SetParamOperations: boolean;
var
 S : String;
begin
 Result:=false;
 if not Assigned(iEParameterized) or (index<0) then exit;
 S:=Text;
 if not iEParameterized.pCheckParameters(index, S, formatSettings) then exit;
 if not iEParameterized.pSetParameters(index, S, formatSettings) then exit;
 Result:=true;
end;

end.

