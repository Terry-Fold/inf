program Project1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {MainForm},
  NewTask in 'NewTask.pas' {Form2},
  StartTask in 'StartTask.pas' {StartTaskform},
  AboutUnit in 'AboutUnit.pas' {AboutForm},
  Settings in 'Settings.pas' {Form3};
  {$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TStartTaskForm, StartTaskForm);
  Application.CreateForm(TAboutForm, AboutForm);
  Application.Run;
end.
