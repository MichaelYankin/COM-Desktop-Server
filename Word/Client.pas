unit Client;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    eDocName: TEdit;
    Label1: TLabel;
    bOpenDoc: TButton;
    bCreateDoc: TButton;
    bSaveDoc: TButton;
    bCloseDoc: TButton;
    bConnect: TButton;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    Label8: TLabel;
    Label9: TLabel;
    bInsertPictureEOF: TButton;
    cbOpen: TCheckBox;
    Label2: TLabel;
    Label3: TLabel;
    bConvertTableToText: TButton;
    Label4: TLabel;
    Label5: TLabel;
    bDeleteWord: TButton;
    eLineToFind: TEdit;
    Label7: TLabel;
    cbDeleteAll: TCheckBox;
    Label6: TLabel;
    bCreateGenericTable: TButton;
    bInsertPicture: TButton;
    bFindServer: TButton;
    procedure bFindServerClick(Sender: TObject);
    procedure bInsertPictureEOFClick(Sender: TObject);
    procedure bConvertTableToTextClick(Sender: TObject);
    procedure bCreateGenericTableClick(Sender: TObject);
    procedure bDeleteWordClick(Sender: TObject);
    procedure bCloseDocClick(Sender: TObject);
    procedure bSaveDocClick(Sender: TObject);
    procedure bOpenDocClick(Sender: TObject);
    procedure bCreateDocClick(Sender: TObject);
    procedure bConnectClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1:      TForm1;
  WB:         OleVariant;
  connected:  boolean;

implementation

{$R *.dfm}

procedure TForm1.bConnectClick(Sender: TObject);
begin
  if connected then
  begin
    WB := Unassigned;
    bConnect.Caption := '������������';
    connected := false;
  end
  else
  begin
    WB := CreateOleObject('Word.Basic');
    bConnect.Caption := '�����������';
    connected := true;
    ShowMessage('����������� �������.');
  end;
  bConnect.Refresh;
end;

// ����������� � ������� ��������� �����
procedure TForm1.bFindServerClick(Sender: TObject);
begin
  try
    WB := GetActiveOleObject('Word.Basic');
    connected := true;
    bConnect.Caption := '�����������';
    bConnect.Refresh;
    ShowMessage('����������� �������.');
  except
    ShowMessage('�� ������� �������� ����������.');
  end;
end;

procedure TForm1.bCreateDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  begin
    WB.FileNew('Normal');
    WB.Insert('�������� ������.'#10#13);
    if cbOpen.Checked then WB.AppShow;
  end;
end;

procedure TForm1.bOpenDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  try
    WB.FileOpen(GetCurrentDir + '\' + eDocName.Text + '.doc');
    WB.AppShow;
  except
    ShowMessage('������ �������� �����!');
  end;
end;


procedure TForm1.bSaveDocClick(Sender: TObject);
var
  fileName: string;
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  begin
    if eDocName.Text <> '' then
    begin
      fileName := GetCurrentDir + '\' + eDocName.Text + '.doc';
      try
        WB.FileSaveAs(fileName, 1);
      except
        ShowMessage('�� ������� ��������� ����!');
      end;
      ShowMessage('���� ������� ��������.');
    end
    else
      ShowMessage('������� ��� �����!');
  end;
end;

procedure TForm1.bInsertPictureEOFClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  begin
    WB.EndOfDocument;
    WB.InsertPicture(ExpandFileName('mpei.jpg'), false, false);
  end;
end;

procedure TForm1.bCloseDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  WB.FileExit(1);
end;

procedure TForm1.bDeleteWordClick(Sender: TObject);
var
  textToFind: string;
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  if eLineToFind.Text = '' then
    ShowMessage('������� ������!')
  else
  begin
    WB.StartOfDocument;
    textToFind := eLineToFind.Text;
    WB.EditFind(textToFind);
    if not WB.EditFindFound then
      ShowMessage('������ �� �������!')
    else
    if cbDeleteAll.Checked then
    repeat
      WB.DeleteWord;
      WB.EditFind(textToFind);
    until not WB.EditFindFound
    else
      WB.DeleteWord;
  end; 
end;

procedure TForm1.bCreateGenericTableClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  begin
    WB.EndOfDocument;
    WB.TableInsertTable(3, 3, Format:=16);
    
    ShowMessage('������� �������. ��������� ��!');
  end;
end;

procedure TForm1.bConvertTableToTextClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('������� ������������ � �������!')
  else
  WB.TableToText;
end;

end.
