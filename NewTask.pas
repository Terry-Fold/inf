unit NewTask;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,
  Vcl.DBCtrls, Vcl.Mask, FileCtrl, IOUtils, Unit1;

type
  TForm2 = class(TForm)
    OkButton: TSpeedButton;
    CancelBotton: TSpeedButton;
    Panel1: TPanel;
    OpenDialog: TOpenDialog;
    GroupBox1: TGroupBox;
    MonthAtCombo: TComboBox;
    MonthToCombo: TComboBox;
    monthdchecklabel: TLabel;
    MonthDiapCheck: TCheckBox;
    MonthLabel: TLabel;
    MonthCombo: TComboBox;
    DChkLabel: TLabel;
    GroupBox2: TGroupBox;
    TfomsButton: TButton;
    TFOMSLabel: TLabel;
    TFOMSFile: TEdit;
    GroupBox3: TGroupBox;
    TypeLabel: TLabel;
    TypeComboBox: TComboBox;
    NameLabel: TLabel;
    NameEdit: TEdit;
    Label1: TLabel;
    tfomsmemo: TMemo;
    procedure MonthDiapCheckClick(Sender: TObject);
    procedure CancelBottonClick(Sender: TObject);
    procedure OkButtonClick(Sender: TObject);
    function ValueCheck:boolean;
    procedure FormShow(Sender: TObject);
    procedure TfomsButtonClick(Sender: TObject);
    function IsDublInf(InfName:string):boolean;
    function IsDublTypeInf(InfType:string):boolean;
    function CpTfomsFile(tfomsfpath:string;InfName:string):boolean;
    procedure CreateNewTask;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  toEdit:boolean;

implementation

{$R *.dfm}

uses StartTask;


procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
mainform.Enabled:=true;
end;

procedure TForm2.FormShow(Sender: TObject);
var i:integer;
begin
// not ToEdit then
 //begin
  nameedit.Clear;
  Typecombobox.Clear;
  TFOMSFile.Clear;
  MonthCombo.Clear;
  MonthAtCombo.Clear;
  MonthToCombo.Clear;
  MonthDiapCheck.Checked:=false;
  NameLabel.Font.Color:=clblack;
  Typelabel.Font.Color:=clblack;
  TFOMSLabel.Font.Color:=clblack;
  MonthLabel.Font.Color:=clblack;
  MonthDiapCheck.Font.Color:=clBlack;
    for i := 1 to 12 do
     begin
      MonthCombo.Items.Add(inttostr(i));
      MonthAtCombo.Items.Add(inttostr(i));
      MonthToCombo.Items.Add(inttostr(i));
     end;
 //end
  //else begin
   typecombobox.Items.Add('Диспансеризация');
   typecombobox.Items.Add('Профосмотр');
   typecombobox.Items.Add('Диспансерное наблюдение');
 // end;
end;

procedure TForm2.MonthDiapCheckClick(Sender: TObject);
begin
 MonthAtCombo.Enabled:=MonthDiapCheck.Checked;
 MonthToCombo.Enabled:=MonthDiapCheck.Checked;
 MonthCombo.Enabled:= not MonthDiapCheck.Checked;
end;

function TForm2.ValueCheck:boolean;
var
  i: Integer;
  monthcheck,monthDcheck:boolean;
begin
 //Проверка заполнения полей
result:=true;
monthcheck:=false;
monthDCheck:=false;
//Имя оповещения
Nameedit.Text := StringReplace(Nameedit.Text , '\','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '/','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , ':','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '*','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '?','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '"','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '<','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '>','', [rfReplaceAll, rfIgnoreCase]);
Nameedit.Text := StringReplace(Nameedit.Text , '|','', [rfReplaceAll, rfIgnoreCase]);
 if  Nameedit.Text='' then
  begin
    namelabel.Font.Color:=clred;
    result:=false;
  end
    else namelabel.Font.Color:=clblack;
 //Вид оповещения
 if (TypeCombobox.Text<>'Диспансеризация')
  and (TypeCombobox.Text<>'Профосмотр')
  and (TypeCombobox.Text<>'Диспансерное наблюдение')
    then
      begin
        TypeLabel.Font.Color:=clred;
        result:=false;
      end
    else typelabel.Font.Color:=clblack;
 //Проверка месяца
 for i := 1 to 12 do
   if monthcombo.Text = inttostr(i) then monthcheck:=true;

 for i := 1 to 12 do
   if (monthAtcombo.Text = inttostr(i)) and (monthTocombo.Text = inttostr(i))
    then monthDcheck:=true;

  if (not monthcheck) and (not monthdiapcheck.Checked) then
    begin
      monthlabel.Font.Color:=clred;
      result:=false;
    end
  else monthlabel.Font.Color:=clblack;


 if monthdiapcheck.Checked then
     begin
       if not monthdcheck then
         begin
           monthdchecklabel.Font.Color:=clred;
           result:=false;
         end
        else
          monthdchecklabel.Font.Color:=clblack;
      try
       if strtoint(Monthatcombo.Text)>strtoint(monthtocombo.Text) then
         begin
           DChkLabel.Visible:=true;
           monthdchecklabel.Font.Color:=clred;
           result:=false;
         end
       else
         begin
           monthdchecklabel.Font.Color:=clblack;
           DChkLabel.Visible:=false;
           result:=true;
         end;
      except end;
     end
     else
      begin
        monthdchecklabel.Font.Color:=clblack;
        DChkLabel.Visible:=false;
      end;

if not FileExists(TFOMSFile.Text) then
    begin
      TFomslabel.Font.Color:=clred;
      result:=false;
    end
  else Tfomslabel.Font.Color:=clblack;


  {
if TFOMSFile.Text=SvodFile.Text then
   begin
    Svodlabel.Font.Color:=clred;
    TFomslabel.Font.Color:=clred;
    result:=false;
   end
 else
  begin
    Tfomslabel.Font.Color:=clblack;
    Svodlabel.Font.Color:=clblack;
  end;
   }
end;

function TForm2.IsDublInf(InfName:string):boolean;
begin
try
with StartTaskForm do begin
   ADOQueryStart.close;
   ADOQueryStart.SQL.Clear;
   ADOQueryStart.SQL.Text:='select * from infs where infs.[name]='''+InfName+'''';
   ADOQueryStart.open;
    if InfName=trimright(adoquerystart.FieldByName('name').AsString) then result := true
     else result := false;
   ADOQueryStart.close;
 end;
except result := false; end;
end;


function TForm2.IsDublTypeInf(InfType:string):boolean;
begin
try
with StartTaskForm do begin
   ADOQueryStart.close;
   ADOQueryStart.SQL.Clear;
   ADOQueryStart.SQL.Text:='select * from infs where infs.[type]='''+InfType+'''';
   ADOQueryStart.open;
    if InfType=trimright(adoquerystart.FieldByName('type').AsString) then
        begin
         if TrimRight(adoquerystart.FieldByName('status_code').AsString)<>'301' then result := true
            else  result := false;
        end
     else result := false;
   ADOQueryStart.close;
 end;
except result := false; end;
end;


procedure TForm2.CreateNewTask;
begin
   with mainform do
            begin
             ADOStoredProc.ProcedureName:='new_Task';
             ADOStoredProc.Parameters.Refresh;
             ADOStoredProc.Parameters.ParamByName('@Inf_id').Value:=Random(100000);
             ADOStoredProc.Parameters.ParamByName('@name').Value:=NameEdit.Text;
             ADOStoredProc.Parameters.ParamByName('@type').Value:=TypeCombobox.Text;
             ADOStoredProc.Parameters.ParamByName('@up_date').Value:=now;
             ADOStoredProc.Parameters.ParamByName('@status_code').Value:= 101;
             ADOStoredProc.Parameters.ParamByName('@Tfoms_path').Value:=tfomsfile.Text;
              if not monthdiapcheck.Checked then
               begin
                ADOStoredProc.Parameters.ParamByName('@MounthAt').Value:= MonthCombo.Text;
                ADOStoredProc.Parameters.ParamByName('@MounthTo').Value:= MonthCombo.Text;
               end
                else
                  begin
                    ADOStoredProc.Parameters.ParamByName('@MounthAt').Value:= MonthAtCombo.Text;
                    ADOStoredProc.Parameters.ParamByName('@MounthTo').Value:= MonthToCombo.Text;
                  end;
               ADOStoredProc.ExecProc;
               AdOQuery.Close;
               ADOQuery.Open;
            end;
end;

function TForm2.CpTfomsFile(tfomsfpath:string;InfName:string):boolean;
var pth:string;
begin
with StartTaskForm do begin
   ADOQueryStart.close;
   ADOQueryStart.SQL.Clear;
   ADOQueryStart.SQL.Text:='declare @pth varchar(100)=(select useNetWD from settings) SELECT CASE @pth When 0 Then WDpth When 1 Then WDNetpth end as pth from settings';
   ADOQueryStart.open;
   pth:=adoquerystart.FieldByName('pth').AsString;

 end;
  TDirectory.CreateDirectory(pth+'\Оповещения\'+InfName);

  if not CopyFile(PChar(tfomsfpath), PChar(pth+'\Оповещения\'+InfName+'\'+extractfilename(tfomsfpath)), false)
   then  messagedlg( 'Невозможно скопировать файл "'+extractfilename(tfomsfpath)+'" в рабочуу папку. '+SysErrorMessage(GetLastError) , mtError, [mbOk], 0, mbOk);

end;

procedure TForm2.OkButtonClick(Sender: TObject);
begin
  if not ValueCheck then messagedlg('Поля не заполнены, или заполнены неверно!' , mtError, [mbOk], 0, mbOk)
    else
      begin
      if IsDublTypeInf(typecombobox.Text) then messagedlg('Перед началом нового информирования '+typecombobox.Text+' необходимо прошлое перенести в архив!' , mtError, [mbOk], 0, mbOk)
        else begin
       if not IsDublInf (NameEdit.Text) then begin
         CpTfomsFile(Tfomsfile.Text,nameedit.Text);
         CreateNewTask;
         form2.Close;
         MainForm.StatusBar.Panels[0].Text:='Новое оповещение с именем '+nameedit.Text+' cоздано успешно!';
        end else messagedlg( 'Оповещение с таким именем уже существует!' , mtError, [mbOk], 0, mbOk);
       end;
      end;
end;


procedure TForm2.TfomsButtonClick(Sender: TObject);
begin
if opendialog.Execute then TFOMSFile.Text:=opendialog.FileName;
end;



procedure TForm2.CancelBottonClick(Sender: TObject);
begin
form2.Close;
end;

end.
