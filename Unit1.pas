unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.DBCGrids, Vcl.WinXPanels, Data.DB, Vcl.Grids, Vcl.DBGrids, Data.Win.ADODB,
  ADOQuery1, ADOConnection1, Vcl.Buttons, ADOStoredProc1, Vcl.StdCtrls,
  Vcl.ValEdit, Vcl.DBCtrls, Registry,ComObj,UITypes;

type
  TMainForm = class(TForm)
    MainPanel: TPanel;
    StatusBar: TStatusBar;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    ADOConnection: TADOConnection1;
    ADOQuery: TADOQuery1;
    DataSource: TDataSource;
    N9: TMenuItem;
    TopPanel: TPanel;
    N10: TMenuItem;
    N11: TMenuItem;
    TableButton: TSpeedButton;
    ControlPanelButton: TSpeedButton;
    ADOStoredProc: TADOStoredProc1;
    N12: TMenuItem;
    MainPanel1: TPanel;
    ViewNotebook: TNotebook;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    MainDBGrid: TDBGrid;
    N5: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    SettingsButton: TSpeedButton;
    ChangeTaskButton: TSpeedButton;
    StartTaskButton: TSpeedButton;
    NewTeskButton: TSpeedButton;
    SpeedButton1: TSpeedButton;
    InfNameLabel: TDBText;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    Statmemo: TRichEdit;
    InfsStatMemo: TRichEdit;
    InfMemo: TRichEdit;
    ProgressBar: TProgressBar;
    procedure FormCreate(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure TableButtonClick(Sender: TObject);
    procedure ControlPanelButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N13Click(Sender: TObject);
    function ReadConnStr:widestring;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function CheckDBConn:boolean;
    procedure SettingsButtonClick(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    Function CheckTableExist(TName:string):boolean;
    procedure N17Click(Sender: TObject);
    function checksettings:boolean;
    procedure AddColoredLine(ARichEdit: TRichEdit; AText: string; AColor: TColor);
    procedure BDConn;
    procedure ADOQueryFetchProgress(DataSet: TCustomADODataSet; Progress,
      MaxProgress: Integer; var EventStatus: TEventStatus);
      procedure tfomspath2Status;
    procedure ChangeTaskButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }

  end;



var
  MainForm: TMainForm;

implementation

uses NewTask, Settings, StartTask, AboutUnit;

{$R *.dfm}

function TMainForm.CheckTableExist(TName:string):boolean;
begin
with starttaskform do
 begin
   ADOQueryStart.Close;
  // ADOQueryStart.SQL.Clear;
   adoquerystart.SQL.Text:= 'if not exists (select * from sysobjects where name='''+TName+''' and xtype=''U'') select 0 else select 1';
   //ADOQueryStart.SQL.Add();
   ADOQueryStart.Open;
   if ADOQueryStart.Fields[0].AsString='1' then
   result:=True else result:=false;
 end;
end;

function TMainForm.ReadConnStr:widestring;
const
  sKey = 'SOFTWARE\Informing\';
var
  rReg: TRegistry;
begin
  rReg := TRegistry.Create;
 // rReg.Access:=KEY_ALL_ACCESS;
  with rReg do
  begin
    RootKey := HKEY_CURRENT_USER;
    if KeyExists(sKey) then
      begin
        OpenKey(sKey,false);
        result:=ReadString('ConnStr');
     //   Showmessage(result);
        CloseKey;
      end
    else result:='not_';
  end;
 rReg.Free;
end;



procedure TMainForm.SettingsButtonClick(Sender: TObject);
begin
SettingsForm.Showmodal;
end;

procedure TMainForm.SpeedButton1Click(Sender: TObject);
begin
AboutForm.ShowModal;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
SettingsForm.Free;
starttaskform.Free;
form2.Free;
AboutForm.Free;
end;

procedure TMainForm.ChangeTaskButtonClick(Sender: TObject);
begin
   NewTask.Form2.Showmodal;
end;

function TMainForm.CheckDBConn:boolean;
begin
result:=false;
  if ReadConnStr<>'not_' then
    begin
      ADOConnection.Close;
      ADOConnection.ConnectionString:=ReadConnStr;
        try
          ADoConnection.Open;
          ADOQuery.Open;
          result:=true;
        except   end;
    end else
      begin
        ADOConnection.Close;
        try
          ADoConnection.Open;
          ADOQuery.Open;
          result:=true;
        except   end;
      end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var i:integer;
begin
 if CheckDBConn then
   begin
     ADOQuery.Close;
     ADOQuery.Open;
   end;
  //Задать размер ширину колонок
i:=0;
while i <= MainDBGrid.Columns.Count - 3 do
 begin
  MainDBGrid.Columns[i].Width:=150;
  i:=i+1;
 end;
      i:=i+1;
  while i <= MainDBGrid.Columns.Count - 1 do
    begin
      MainDBGrid.Columns[i].Width:=110;
      i:=i+1;
    end;
//
end;

 procedure  TMainForm.AddColoredLine(ARichEdit: TRichEdit; AText: string; AColor: TColor);
 begin
   with ARichEdit do
   begin
     SelStart := Length(Text);
     SelAttributes.Color := AColor;
     SelAttributes.Size := 8;
     SelAttributes.Name := 'MS Sans Serif';
     Lines.Add(AText);
   end;
 end;


function TMainForm.checksettings:boolean;
begin
//Statmemo.Clear;

result:=true;

  settingsform.ADOConnSet.Close;
  settingsform.ADOqueryset.Close;
  settingsform.ADOqueryset.Open;
  settingsform.ADOConnSet.Open;

if settingsform.adoqueryset.FieldByName('WDpth').AsString='' then
  begin
    AddColoredLine(statmemo,'Не указана рабочая папка!', clRed);
    result:=false;
  end
  else AddColoredLine(statmemo,'Рабочая папка найдена!', clGreen);

if settingsform.adoqueryset.FieldByName('SMSDispTxt').AsString='' then
  begin
    AddColoredLine(statmemo,'Не указан текст СМС для Диспансеризации!', clRed);
    result:=false;
  end
    else AddColoredLine(statmemo,'Текст СМС для Диспансеризации найден!', clGreen);

if settingsform.adoqueryset.FieldByName('SMSProfTxt').AsString='' then
  begin
   AddColoredLine(statmemo,'Не указан текст СМС для Профосмотра!', clRed);
   result:=false;
  end
    else AddColoredLine(statmemo,'Текст СМС для Профосмотра найден!', clGreen);

if settingsform.adoqueryset.FieldByName('SMSDispNabTxt').AsString='' then
  begin
    AddColoredLine(statmemo,'Не указан текст СМС для Диспансерного наблюдения!', clRed);
    result:=false;
  end
    else AddColoredLine(statmemo,'Текст СМС для Диспансерного наблюдения найден!', clGreen);

if settingsform.adoqueryset.FieldByName('ViberDispTxt').AsString='' then
  begin
   AddColoredLine(statmemo,'Не указан текст Viber для Диспансеризации!', clRed);
   result:=false;
  end
    else AddColoredLine(statmemo,'Текст Viber для Диспансеризации найден!', clGreen);

if settingsform.adoqueryset.FieldByName('ViberProfTxt').AsString='' then
  begin
    AddColoredLine(statmemo,'Не указан текст Viber для Профосмотра!', clRed);
    result:=false;
  end
    else AddColoredLine(statmemo,'Текст Viber для Профосмотра найден!', clGreen);

if settingsform.adoqueryset.FieldByName('ViberDispNabTxt').AsString='' then
  begin
   AddColoredLine(statmemo,'Не указан текст Viber для Диспансерного наблюдения!', clRed);
   result:=false;
  end
    else AddColoredLine(statmemo,'Текст Viber для Диспансерного наблюдения найден!', clGreen);

if settingsform.adoqueryset.FieldByName('Svdpth').AsString='' then
  begin
    AddColoredLine(statmemo,'Не указан файл "Свод маршрутизации"!', clRed);
    result:=false;
  end
    else AddColoredLine(statmemo,'Файл "Свод маршрутизации" указан!', clGreen);

if not CheckTableExist('SvodMarsh') then
  begin
   AddColoredLine(statmemo,'Таблица "Свод маршрутизации" не существует!', clRed);
   result:=false;
  end
    else  AddColoredLine(statmemo,'Таблица "Свод маршрутизации" существует!', clGreen);

if settingsform.adoqueryset.FieldByName('useNetWD').AsString='1' then
  begin
    if settingsform.adoqueryset.FieldByName('WDNetpth').AsString = '' then
      begin
        AddColoredLine(statmemo,'Не указан сетевой путь к рабочей папке!', clRed);
        result:=false;
      end
      else AddColoredLine(statmemo,'Сетевой путь к рабочей папке найден!', clGreen);
  end;
end;

procedure TMainForm.ADOQueryFetchProgress(DataSet: TCustomADODataSet; Progress,
  MaxProgress: Integer; var EventStatus: TEventStatus);
begin
  mainform.ProgressBar.Max:=MaxProgress;
  application.ProcessMessages;
  mainform.ProgressBar.Position:= Progress;
end;

procedure TMainForm.BDConn;
begin
 if CheckDBConn then AddColoredLine(statmemo,'Соединение с БД - УСПЕШНО!', clGreen)
  else begin
    infmemo.Clear;
    infsstatmemo.Lines.Add('Невозможно получить статистику!');
    AddColoredLine(statmemo,'Нет связи с БД! ОШИБКА!', clRed);
    infmemo.Lines.Add('Нет данных!');
  end;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
Statmemo.Clear;
 //Проверка соединения с БД
    BDConn;
  //Проверка настроек
     checksettings;
 //Статус
 tfomspath2Status;
end;

function get_shortname(fname:string):string;
begin
  while pos('\',fname)>0 do
    fname:=copy(fname,pos('\',fname)+1,length(fname));
  result:=fname;
end;

procedure TMainForm.tfomspath2Status;
begin
   with starttaskform do
  begin
    ADOQueryStart.Close;
    ADOQueryStart.SQL.Text:='select * from Infs';
    ADOQueryStart.Open;

    statusbar.Panels[0].Text:= get_shortname(adoquerystart.FieldByName('Tfoms_path').AsString);

    // infNamewrklabel.Caption;

  end;
end;


procedure TMainForm.N12Click(Sender: TObject);
begin
StartTaskForm.show;
end;

procedure TMainForm.N13Click(Sender: TObject);
begin
AboutForm.showmodal;
end;

procedure TMainForm.N16Click(Sender: TObject);
begin
SettingsForm.SetTabNB.PageIndex:=2;
SettingsForm.Showmodal;
end;

procedure TMainForm.N17Click(Sender: TObject);
begin
mainform.Close;
end;

procedure TMainForm.N5Click(Sender: TObject);
begin
SettingsForm.SetTabNB.PageIndex:=0;
SettingsForm.Showmodal;
end;

procedure TMainForm.N8Click(Sender: TObject);
begin
 Statmemo.Clear;
  BDConn;
if not checksettings then
  begin
    viewnotebook.PageIndex:=0;
    messagedlg( 'Невозможно создать новое информирование пока не указаны все настройки.' , mtError,  [mbOk], 0);
  end else begin
     mainform.Enabled:=false;
    NewTask.Form2.Showmodal;
  end;
end;

procedure TMainForm.N9Click(Sender: TObject);
begin
SettingsForm.SetTabNB.PageIndex:=1;
SettingsForm.Showmodal;
end;

procedure TMainForm.ControlPanelButtonClick(Sender: TObject);
begin
ViewNotebook.PageIndex:=0;
end;

procedure TMainForm.TableButtonClick(Sender: TObject);
begin
ViewNotebook.PageIndex:=1;
end;

end.
