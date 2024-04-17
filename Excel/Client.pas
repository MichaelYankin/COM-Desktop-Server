unit Client;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Excel_TLB;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    eDocName: TEdit;
    bCreateDoc: TButton;
    bSaveDoc: TButton;
    bOpenDoc: TButton;
    bCloseDoc: TButton;
    cbOpen: TCheckBox;
    bConnect: TButton;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    ePrecisionValue: TEdit;
    bSetPrecision: TButton;
    GroupBox3: TGroupBox;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    bSetHeightAndWidth: TButton;
    Label13: TLabel;
    GroupBox4: TGroupBox;
    bAddChartSource: TButton;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label14: TLabel;
    procedure FormDestroy(Sender: TObject);
    procedure bAddChartSourceClick(Sender: TObject);
    procedure bSetHeightAndWidthClick(Sender: TObject);
    procedure bSetPrecisionClick(Sender: TObject);
    procedure bSaveDocClick(Sender: TObject);
    procedure bCloseDocClick(Sender: TObject);
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
  EA:         _ApplicationDisp;
  EWB:        WorkBooksDisp;
  connected:  boolean;

implementation

{$R *.dfm}

procedure TForm1.bCloseDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
    EWB.Close;
end;

procedure TForm1.bConnectClick(Sender: TObject);
begin
  if not connected then
  begin
    EA := CoExcelApplication.Create As _ApplicationDisp;
    EWB := EA.WorkBooks As WorkBooksDisp;
    connected := true;
    bConnect.Caption := 'Отключиться';
    ShowMessage('Подключение успешно.');
  end
  else
  begin
    EA.Quit;
    if Assigned(EA) then
     EA := nil;
    EWB := nil;
    connected := false;
    bConnect.Caption := 'Подключиться';
  end;
  bConnect.Refresh;
end;

procedure TForm1.bCreateDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
  begin
    EWB.Add(EmptyParam);
    EA.Visible := cbOpen.Checked;
  end;
end;

procedure TForm1.bOpenDocClick(Sender: TObject);
var
  fileName: string;
begin
  if eDocName.Text = '' then
    ShowMessage('Введите название файла!')
  else
  begin
  try
    fileName := GetCurrentDir + '\' + eDocName.Text + '.xls';
    EA := CoExcelApplication.Create As _ApplicationDisp;
    EWB := EA.WorkBooks As WorkBooksDisp;
    EWB.Open(fileName, EmptyParam, false, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
      EmptyParam, true, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    EWB.Item[1].Activate;
    EA.Visible := True;
    connected := true;
    bConnect.Caption := 'Отключиться';
    bConnect.Refresh;
  except
    ShowMessage('Файл с указанным именем не найден!');
  end;
  end;
end;

procedure TForm1.bSaveDocClick(Sender: TObject);
begin
  if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
  if eDocName.Text = '' then
    ShowMessage('Введите название файла!')
  else
  begin
    (EA.ActiveWorkbook As _WorkBookDisp).SaveCopyAs(GetCurrentDir + '\' + eDocName.Text + '.xls');
    if (EA.ActiveWorkbook As _WorkBookDisp).Saved then
      ShowMessage('Файл сохранен.')
    else
      ShowMessage('Ошибка сохранения файла!');
  end;
end;

procedure TForm1.bSetPrecisionClick(Sender: TObject);
begin
   if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
  begin
    try
      EA.FixedDecimal := true;
      EA.FixedDecimalPlaces := StrToInt(ePrecisionValue.Text);
    except
      on EConvertError do
        ShowMessage('Введите целое число!');
    end;
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  if Assigned(EA) then
    EA.Quit;
end;

procedure TForm1.bAddChartSourceClick(Sender: TObject);
var
  EWS:      _WorkSheetDisp;
  myChart:  _ChartDisp;
begin
  if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
  begin
  try
    EWS := EA.ActiveSheet as _WorkSheetDisp;
    (EA.ActiveWorkbook As _Workbook).Charts.Add(EmptyParam, EmptyParam, 1, EmptyParam, 0);
    myChart := EA.ActiveChart As _ChartDisp;
    myChart.ChartType := xlLine;
    myChart.SetSourceData(EWS.Range['A1', 'C3'], xlRows);
    myChart.Location(xlLocationAsObject, 'Лист1');
  except
    ShowMessage('В диапазоне содержатся неверные данные!');
  end;
  end;
end;

procedure TForm1.bSetHeightAndWidthClick(Sender: TObject);
var
  EWS: _WorkSheetDisp;
begin
  if not connected then
    ShowMessage('Сначала подключитесь к серверу!')
  else
  begin
    EWS := EA.ActiveSheet as _WorkSheetDisp;
    EWS.Rows.RowHeight := EWS.StandardHeight;
    EWS.Columns.ColumnWidth := EWS.StandardWidth;
  end;
end;

end.
