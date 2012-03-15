program HattrickKoopjesScanner;

uses
  Forms,
  formHattrickKoopjesScanner in 'formHattrickKoopjesScanner.pas' {frmHattrickKoopjesScanner},
  uHattrick in '..\Hattrick Scanner\uHattrick.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmHattrickKoopjesScanner, frmHattrickKoopjesScanner);
  Application.Run;
end.
