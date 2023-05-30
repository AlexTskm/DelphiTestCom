{$ALIGN OFF}
{$H+}
unit MyInterface;

interface

uses System.SysUtils, Activex;

const
// ���� � ���������� � �������
 ThePathToTheApplication = 'SOFTWARE\AlfaLab\Test\';

// ������ ���������� ����
Class_Of_TIncTheDate:TCLSID=
'{12345678-0001-AADD-0001-080520232013}';

// ������ ���������� ����������� �����
Class_Of_TSquareExtraction:TCLSID=
'{12345678-0001-AADD-0001-110520230844}';


type
// ���� ������

// ������ ����������������� ��������
PRecordingParameters =^TRecordingParameters;
TRecordingParameters = record
// �������� ���������
  pNameParameter : AnsiString;
// ����������� ���������� ���������
  pHintParameter : AnsiString;
// ��� ���������(������������/��������������)
  pRequired      : boolean;
// ��� ��������� ������������ �������� VarType(������������� �����, ������� �������� �� ������� � �.�)
  pEnumeration   : Variant;
// ��������
  pParameters    : Variant;
end;

// ��������� IParameterizedOperations
// �������� ���������, ������� ����������������� ��������
IParameterizedOperations = Interface(IUnknown)
['{12345678-0001-AADD-0000-060520231408}']
// ����������� ������� ����������������� ��������
// ��������� ������� ����������������� �������� (����������: <>0 ������)
  function  poCheck:longint;StdCall;
// ��������� ������� ����������������� �������� (����������: <>0 ������)
  function  poExecute:longint;StdCall;
// �������� ��������� ������� ����������������� ��������
  procedure  poClearParameters;StdCall;
// ������� ����� ��������� ������ poExecute
  function poGetResult: Variant;StdCall;

// ����� �������� ����������������� ��������
  function poGetNameParameters:AnsiString;Stdcall;
// ���������� ���������� ������� ����������������� �������� (����������: < 0 ������)
  function poCountOfParameters:longint;StdCall;
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
// ��������� �������� �� ������ pVerify ��� ��������� Index ���������� � 1 . � ������ ������ ���������� ���������� pVerify
  function pCheckParameters(Index:longint; pVerify : AnsiString; const Formatting : TFormatSettings) : boolean;StdCall;
end;


implementation

end.
