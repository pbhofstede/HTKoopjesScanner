unit formHattrickKoopjesScanner;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  OleCtrls, SHDocVw_EWB, EwbCore, EmbeddedWB, ExtCtrls, dxPageControl, Db,
  IBDatabase, dxDBTLCl, dxGrClms, dxTL, dxDBCtrl, dxDBGrid, uHattrick,
  IBCustomDataSet, dxCntner, synDBGrid, dxBar, ImgList, StdCtrls, IBSQL,
  TntStdCtrls, dxBarExtItems, cxClasses;

type
  TfrmHattrickKoopjesScanner = class(TForm)
    pgctrlLijst: TdxPageControl;
    tsBidwar: TdxTabSheet;
    pnlLeft: TPanel;
    HTBrowser: TEmbeddedWB;
    ibdbHTInfo: TIBDatabase;
    ibtrMain: TIBTransaction;
    dxBarManager1: TdxBarManager;
    pmKoopjes: TdxBarPopupMenu;
    btnOpenSpeler: TdxBarButton;
    dxBarDockControl1: TdxBarDockControl;
    btnStart: TdxBarButton;
    ImageList1: TImageList;
    tmrRefresh: TTimer;
    vMemo: TTntMemo;
    btnStop: TdxBarButton;
    lblStatus: TdxBarStatic;
    tmrTimer: TTimer;
    mmLog: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure HTBrowserBeforeNavigate2(ASender: TObject;
      const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
      Headers: OleVariant; var Cancel: WordBool);
    procedure HTBrowserDocumentComplete(ASender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);
    procedure HTBrowserDownloadBegin(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure tmrRefreshTimer(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure tmrTimerTimer(Sender: TObject);
  private
    { Private declarations }
    FDocumentCompleted: Boolean;
    FURL: String;
    FUserName: String;
    FPassWord: String;
    FKoopjesMarge: integer;
    FKoopjesMargePerc: integer;
    FKoopjesLoonWeken: integer;
    FTeamID: integer;
    FTPEMarge: integer;
    FVorigeClubs: double;
    FMakelaarsPerc: double;
    FFirstScanTijdstip: TDateTime;
    FTransferBudget: integer;

    procedure LogIn;
    procedure BrowseTo(aBrowser: TEmbeddedWB; aURL: String);
    function BrowserClick(aBrowser:TEmbeddedWB;const aBrowserObjectNames:array of String;aShowError:boolean=TRUE):boolean;
    procedure BrowseToPlayer(aPlayerID: integer);
    procedure Refresh;
    procedure StreamToMemo(aBrowser: TEmbeddedWB);
    procedure AddLog(aString: String);
    procedure ScanSpelers(aDeadline, aMinLeeftijd, aMaxLeeftijd, aSkill, aMinSkill, aMaxSkill: integer);
    function BrowseToLink(aBrowser: TEmbeddedWB; aLink: String): boolean;
    procedure ScanKoopjes;
    function BrowseToNextPage(aPage: integer): boolean;
    function ParsePrijzen(var vPlayers: integer): double;
    function BerekenMaxPrijs(aTransferprijs: double; aLoon, aLeeftijd,
      aBlessureWeken: integer): double;
    function ParsePlayerInfo: TTSISet;
    procedure SaveKoopje(aScoutingID, aPlayerID, aTPE, aHoogsteBod,
      aMaxBod: integer; aDeadline: TDateTime; vTPEAantalSpelers: integer);
    function SaveScouting(aTSISet: TTSISet;
      aTalentScouting: boolean): integer;
  public
    { Public declarations } 
    property DocumentCompleted:boolean read FDocumentCompleted write FDocumentCompleted;
    property TransferBudget: integer read FTransferBudget write FTransferBudget;
  end;

CONST
  KEEPEN = 1;
  CONDITIE = 2;
  VERDEDIGEN = 3;
  POSITIESPEL = 4;
  VLEUGELSPEL = 5;
  SCOREN = 6;
  SPELHERVATTING = 7;
  PASSEN = 8;

var
  frmHattrickKoopjesScanner: TfrmHattrickKoopjesScanner;

implementation
uses
  inifiles, esbDates, Math, uBibConv, ActiveX, uBibMath, uBibDB, uBibString;

{$R *.DFM}

{ TForm4 }

procedure TfrmHattrickKoopjesScanner.LogIn;
begin
  BrowseTo(HTBrowser,'http://www.hattrick.org');
  if (FUserName <> '') then
  begin
    if (FUserName <> '') then
    begin
      SetValue(HTBrowser,['ctl00$ucSubMenu$txtUserName','ctl00$ucSubMenu$ucLogin$txtUserName',
        'ctl00_ctl00_CPContent_ucSubMenu_ucLogin_txtUserName'], FUserName);
      SetValue(HTBrowser,['ctl00$ucSubMenu$txtPassword', 'ctl00$ucSubMenu$ucLogin$txtPassword',
        'ctl00_ctl00_CPContent_ucSubMenu_ucLogin_txtPassword'], FPassWord);
      BrowserClick(HTBrowser, ['ctl00$ucSubMenu$butLogin', 'ctl00$ucSubMenu$ucLogin$butLogin',
        'ctl00$ctl00$CPContent$ucSubMenu$ucLogin$butLogin']);
    end;

    AddLog('Login done');
  end;
end;

procedure TfrmHattrickKoopjesScanner.BrowseTo(aBrowser:TEmbeddedWB;aURL:String);
begin
  FDocumentCompleted := FALSE;
  aBrowser.Navigate(aURL);

  while not DocumentCompleted do
  begin
    Application.HandleMessage;
  end;
end;

procedure TfrmHattrickKoopjesScanner.FormCreate(Sender: TObject);
begin
  FDocumentCompleted := FALSE;
  FURL := '';

  with TIniFile.Create(ExtractFilePath(Application.ExeName) +  'HTScanner.ini') do
  begin
    try
      FUserName := ReadString('ALGEMEEN','USERNAME','');
      FPassWord := ReadString('ALGEMEEN','PASSWORD','');
      FTeamID := ReadInteger('ALGEMEEN', 'TEAMID', 0);
      
      ibdbHTInfo.DatabaseName := ReadString('ALGEMEEN','DATABASE',
        Format('localhost:%sDATA\HT_INFO.GDB',
        [ExtractFilePath(Application.ExeName)]));

      FTPEMarge := ReadInteger('KOOPJES', 'TPEMARGE', 5);
      FMakelaarsPerc := ReadFloat('KOOPJES', 'MAKELAARSPERC', 7.55);
      FVorigeClubs := ReadFloat('KOOPJES', 'VORIGECLUBS', 4.5);
      FKoopjesMarge := ReadInteger('KOOPJES', 'MARGE', 150000);
      FKoopjesMargePerc := ReadInteger('KOOPJES', 'MARGEPERC', 10);
      FKoopjesLoonWeken := ReadInteger('KOOPJES', 'LOONWEKEN', 3);
    finally
      Free;
    end;
  end;

  TransferBudget := 6800000;

  if (ibdbHTInfo.DatabaseName <> '') then
  begin  
    ibdbHTInfo.Open;
  end;
end;

function TfrmHattrickKoopjesScanner.BrowserClick(aBrowser:TEmbeddedWB;const aBrowserObjectNames:array of String
  ;aShowError:boolean=TRUE):boolean;
var
  vItem:Variant;
  vTime:TDateTime;
  i,
  vCount: integer;
begin
  result := FALSE;
  vItem := Unassigned;
  i := 0;
  while VarIsEmpty(vItem) and (i <= High(aBrowserObjectNames)) do
  begin
    vItem := uHattrick.GetBrowserObject(aBrowser,aBrowserObjectNames[i]);
    inc(i);
  end;
  
  if VarIsEmpty(vItem)  then
  begin
    if (aShowError) then
    //ShowMessage('aBrowserObjectName en aBrowserObjectName2 bestaan niet! '+aBrowserObjectName+ ' | ' + aBrowserObjectName2);
  end
  else
  begin
    result := TRUE;
    FDocumentCompleted := FALSE;
    vItem.Click;
    vTime := Now;

    while not DocumentCompleted do
    begin
      Application.HandleMessage;
      if (esbDates.TimeApartInSecs(Now,vTime) > 10) then
      begin
        FDocumentCompleted := TRUE;
        BrowserClick(aBrowser,aBrowserObjectNames);
      end;
    end;
  end;

  for vCount := 0 to 9 do
  begin
    Application.ProcessMessages;
    Sleep(10);
  end;
end;

procedure TfrmHattrickKoopjesScanner.HTBrowserBeforeNavigate2(ASender: TObject;
  const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
  Headers: OleVariant; var Cancel: WordBool);
begin
  DocumentCompleted := FALSE;
end;

procedure TfrmHattrickKoopjesScanner.HTBrowserDocumentComplete(ASender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  if (pDisp = TEmbeddedWB(aSender).DefaultInterface) then
  begin
    DocumentCompleted := TRUE;
    FUrl := URL;

    //Application.ProcessMessages;
  end;
end;

procedure TfrmHattrickKoopjesScanner.HTBrowserDownloadBegin(Sender: TObject);
begin
  DocumentCompleted := FALSE;
end;

procedure TfrmHattrickKoopjesScanner.BrowseToPlayer(aPlayerID: integer);
var
  vURL: String;
begin
  vURL := FURL;
  vURL := Copy(vURL, 1, Pos('.hattrick.org/', FURL));
  vURL := Format('%shattrick.org/Club/Players/Player.aspx?playerId=%d', [vURL, aPlayerID]);

  BrowseTo(HTBrowser, vURL);
end;

procedure TfrmHattrickKoopjesScanner.btnStartClick(Sender: TObject);
begin
  AddLog('Starting bot');
  
  btnStart.Enabled := FALSE;
  btnStop.Enabled := TRUE;
  tmrTimer.Enabled := TRUE;

  tmrRefreshTimer(Sender);
end;

procedure TfrmHattrickKoopjesScanner.Refresh;
var
  vStartTijd: TDateTime;
begin
  vStartTijd := now;
  try
    FFirstScanTijdstip := 0;

    Login;


    ScanSpelers(13, 27, 33, POSITIESPEL, 15, 18);
    ScanSpelers(13, 21, 26, POSITIESPEL, 15, 18);
    ScanSpelers(13, 27, 33, POSITIESPEL, 11, 14);
    ScanSpelers(13, 21, 26, POSITIESPEL, 11, 14);

    ScanSpelers(13, 27, 33, VERDEDIGEN, 15, 18);
    ScanSpelers(13, 21, 26, VERDEDIGEN, 15, 18);
    ScanSpelers(13, 27, 33, VERDEDIGEN, 11, 14);
    ScanSpelers(13, 21, 26, VERDEDIGEN, 11, 14);

    ScanSpelers(13, 27, 33, SCOREN, 15, 18);
    ScanSpelers(13, 21, 26, SCOREN, 15, 18);
    ScanSpelers(13, 27, 33, SCOREN, 11, 14);
    ScanSpelers(13, 21, 26, SCOREN, 11, 14);

    ScanSpelers(13, 27, 33, KEEPEN, 15, 18);
    ScanSpelers(13, 21, 26, KEEPEN, 15, 18);
    ScanSpelers(13, 27, 33, KEEPEN, 11, 14);
    ScanSpelers(13, 21, 26, KEEPEN, 11, 14);

    ScanSpelers(13, 27, 33, VLEUGELSPEL, 15, 18);
    ScanSpelers(13, 21, 26, VLEUGELSPEL, 15, 18);
    ScanSpelers(13, 27, 33, VLEUGELSPEL, 11, 14);
    ScanSpelers(13, 21, 26, VLEUGELSPEL, 11, 14);

  finally
    vStartTijd := vStartTijd + (10 / 24);
    FFirstScanTijdstip := vStartTijd;
    tmrRefresh.Interval := Ceil(ESBDates.TimeApartInSecs(Now, vStartTijd) * 1000);

    BrowserClick(HTBrowser, ['ctl00_ucMenu_hypLogout']);
  end;
end;

procedure TfrmHattrickKoopjesScanner.tmrRefreshTimer(Sender: TObject);
begin
  tmrRefresh.Enabled := FALSE;
  try
    Refresh;
  finally
    tmrRefresh.Enabled := TRUE;
  end;
end;

procedure TfrmHattrickKoopjesScanner.StreamToMemo(aBrowser:TEmbeddedWB);
var
  sm:TMemoryStream;
  sa:IStream;
  vResult : OLEVariant;
begin
  vMemo.Lines.Clear;

  sm := TMemoryStream.Create;
  try
    sa := TStreamAdapter.Create(sm, soReference) as IStream;
    vResult := (HTBrowser.Document as IPersistStreamInit).Save(sa, TRUE);

    sm.Seek(0, soFromBeginning);
    vMemo.Lines.LoadFromStream(sm, FALSE);
  finally
    FreeAndNil(sm);
  end;
end;

procedure TfrmHattrickKoopjesScanner.btnStopClick(Sender: TObject);
begin
  AddLog('Stopping bot');

  tmrTimer.Enabled := FALSE;
  btnStop.Enabled := FALSE;
  btnStart.Enabled := TRUE;
  tmrRefresh.Enabled := FALSE;

  lblStatus.Caption := '';
end;

procedure TfrmHattrickKoopjesScanner.tmrTimerTimer(Sender: TObject);
var
  vMinuten,
  vSeconden,
  vTotSeconden: integer;
  vMin: String;
begin
  if (FFirstScanTijdstip > 0) then
  begin
    vTotSeconden := Ceil(ESBDates.TimeApartInSecs(Now, FFirstScanTijdstip));
    if (vTotSeconden < 0) then
    begin
      vMin := '-';
      vTotSeconden := Abs(vTotSeconden);
    end
    else
    begin
      vMin := '';
    end;

    vMinuten := vTotSeconden div 60;
    vSeconden := vTotSeconden mod 60;
    if (vSeconden < 10) then
    begin
      lblStatus.Caption := Format('Refresh over %s%d:0%d', [vMin, vMinuten, vSeconden]);
    end
    else
    begin
      lblStatus.Caption := Format('Refresh over %s%d:%d', [vMin, vMinuten, vSeconden]);
    end;
  end;
end;

procedure TfrmHattrickKoopjesScanner.AddLog(aString: String);
begin
  mmLog.Lines.Text :=
    Format('%s %s', [FormatDateTime('dd-mm-yyyy hh:nn:ss', Now), aString]) + #13#10 +
    mmLog.Lines.Text;
end;

procedure TfrmHattrickKoopjesScanner.ScanSpelers(aDeadline, aMinLeeftijd, aMaxLeeftijd,
  aSkill, aMinSkill, aMaxSkill: integer);
begin
  AddLog(Format('Scanning %d-%d %d %d - %d', [aMinLeeftijd, aMaxLeeftijd, aSkill, aMinSkill, aMaxSkill]));

  BrowseToLink(HTBrowser,'/World/Transfers');

  SetValue(HTBrowser,['ctl00$CPMain$ddlDeadline'], IntToStr(aDeadline));
  SetValue(HTBrowser,['ctl00$CPMain$ddlAgeMin'], IntToStr(aMinLeeftijd));
  SetValue(HTBrowser,['ctl00$CPMain$ddlAgeMax'], IntToStr(aMaxLeeftijd));

  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill1'], IntToStr(aSkill));
  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill1Min'], IntToStr(aMinSkill));
  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill1Max'], IntToStr(aMaxSkill));

  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill2'], IntToStr(CONDITIE));
  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill2Min'], IntToStr(4));
  SetValue(HTBrowser,['ctl00$CPMain$ddlSkill2Max'], IntToStr(9));

  BrowserClick(HTBrowser, ['ctl00$CPMain$butSearch']);

  ScanKoopjes;
end;

function TfrmHattrickKoopjesScanner.BrowseToLink(aBrowser:TEmbeddedWB; aLink: String):boolean;
var
  vURL:String;
begin
  vUrl := GetLink(aBrowser, aLink, TRUE);

  if (vURL = '') then
  begin
    vUrl := GetLink(aBrowser, aLink, FALSE);
  end;
  
  if (vURL <> '') then
  begin
    BrowseTo(aBrowser,vUrl);
  end;

  result := (vURL <> '');
end;

procedure TfrmHattrickKoopjesScanner.ScanKoopjes;
const
  PLAYERID = '<a href="/Club/Players/Player.aspx?PlayerID=';
  DEADLINE = 'TransferPlayer_lblDeadline';
var
  vPages, vPos, vTempPos, vMaxPages, vCount, vTPEAantalSpelers, vLeeftijd, vGevondenKoopjes:integer;
  vTransferprijsEvaluatie, vMaxPrijs: double;
  vPlayerIDs: array of integer;
  vPlayerPrizes: array of integer;
  vPlayerDeadlines: array of TDateTime;
  vText, vTemp, vTempText, vURL: String;
  vTSISet: TTSISet;
begin
  vGevondenKoopjes := 0;

  uBibDB.ExecSQL(ibdbHTInfo, 'DELETE FROM KOOPJES WHERE (DEADLINE IS NULL) or (DEADLINE < (CURRENT_TIMESTAMP - 0.2))', [], []);

  try
    vMaxPages := 40;
    vPages := 1;

    while vPages <= vMaxPages do
    begin
      lblStatus.Caption := Format('%d/%d',[vPages,vMaxPages]);

      StreamToMemo(HTBrowser);
      vText := Copy(vMemo.Text, Pos('Zoekresultaat', vMemo.Text), Length(vMemo.Text));

      vPos := Pos(PLAYERID, vText);

      while (vPos > 0) do
      begin
        vText := Copy(vText, vPos + Length(PLAYERID), Length(vText));
        vPos := Pos(PLAYERID, vText);
        vTempText := Copy(vText, 1, vPos);

        vTemp := Copy(vTempText, 1, Pos('&amp;', vTempText) - 1);
        SetLength(vPlayerIDs, Length(vPlayerIDs) + 1);
        vPlayerIDs[Length(vPlayerIDs) - 1] := uBibConv.AnyStrToInt(vTemp);

        vTemp := Copy(vTempText, 1, Pos('&nbsp;€', vTempText) - 1);
        vTempPos := uBibString.GetLastPos('<td>', vTemp);
        vTemp := Copy(vTemp, vTempPos + 4, Length(vTemp));

        SetLength(vPlayerPrizes, Length(vPlayerPrizes) + 1);
        vPlayerPrizes[Length(vPlayerPrizes) - 1] := uBibConv.AnyStrToInt(VerwijderSpaties(Trim(vTemp)));

        SetLength(vPlayerDeadlines, Length(vPlayerDeadlines) + 1);
        if (Pos(DEADLINE, vTempText) > 0) then
        begin
          vTemp := Copy(vTempText, Pos(DEADLINE, vTempText) + Length(DEADLINE), Length(vTempText));
          vTemp := Copy(vTemp, Pos('">', vTemp) + 2, Length(vTemp));
          vTemp := Copy(vTemp, 1, Pos('</span>', vTemp) - 1);

          vPlayerDeadlines[Length(vPlayerDeadlines) - 1] := StrToDateTime(Trim(vTemp));
        end;
      end;

      inc(vPages);

      // Pagina verder
      if (vPages <= vMaxPages) then
      begin
        if not BrowseToNextPage(vPages - 1) then
        begin
          vPages := vMaxPages + 1;
        end;
      end;
    end;

    for vCount := 0 to Length(vPlayerIDs) - 1 do
    begin
      lblStatus.Caption := Format('%d/%d',[vCount,Length(vPlayerIDs)]);

      if (uBibDB.GetFieldValue(ibdbHTInfo, 'KOOPJES', ['PLAYER_ID'], [vPlayerIDs[vCount]], 'ID', srtInteger) = 0) then
      begin
        vURL := FURL;
        vURL := Copy(vURL, 1, Pos('.hattrick.org/', FURL));
        vURL := Format('%shattrick.org/Club/Transfers/TransferCompare.aspx?playerId=%d', [vURL, vPlayerIDs[vCount]]);

        BrowseTo(HTBrowser, vURL);

        vTransferprijsEvaluatie := ParsePrijzen(vTPEAantalSpelers);

        vMaxPrijs := BerekenMaxPrijs(vTransferprijsEvaluatie, 0, 0, 0);

        if (vPlayerPrizes[vCount] < vMaxPrijs) and
           (vTPEAantalSpelers > 1) then
        begin
          BrowseToPlayer(vPlayerIDs[vCount]);

          vTSISet := ParsePlayerInfo;
          try
            vLeeftijd := uBibConv.AnyStrToInt(vTSISet.Leeftijd);
            vMaxPrijs := BerekenMaxPrijs(vTransferprijsEvaluatie, vTSISet.Loon, vLeeftijd, vTSISet.WekenBlessure);

            SaveKoopje(SaveScouting(vTSISet, FALSE), vTSISet.PlayerID, Ceil(vTransferprijsEvaluatie),
              Ceil(Max(vTSISet.VraagPrijs, vTSISet.HoogsteBod)),
              Ceil(vMaxPrijs), vTSISet.DeadLine, vTPEAantalSpelers);

            if (Max(vTSISet.VraagPrijs, vTSISet.HoogsteBod) <= vMaxPrijs) then
            begin
              Inc(vGevondenKoopjes);
            end;
          finally
            vTSISet.Free;
          end;
        end
        else
        begin
          SaveKoopje(0, vPlayerIDs[vCount], Ceil(vTransferprijsEvaluatie), vPlayerPrizes[vCount],
                  Ceil(vMaxPrijs), vPlayerDeadlines[vCount], vTPEAantalSpelers);
        end;
      end;
    end;
  finally
    AddLog(Format('%d nieuwe koopjes gescout!',[vGevondenKoopjes]));
  end;
end;

function TfrmHattrickKoopjesScanner.BrowseToNextPage(aPage:integer):boolean;
var
  vElement:String;
  vNextPage: Variant;
  vIndex:integer;
begin
  result := TRUE;
  vIndex := aPage;

  if (aPage > 10) then
  begin
    vIndex := 9 + ((Floor((aPage -1) / 10) - 1) * 10);

    vIndex := aPage - vIndex;
  end;

  vElement := Format('ctl00_CPMain_ucPager_repPages_ctl%.2d_p%d',[vIndex,aPage]);

  vNextPage := HTBrowser.OLEObject.Document.GetElementByID(vElement);

  if (varIsEmpty(vNextPage)) then
  begin
    StreamToMemo(HTBrowser);
    vMemo.Lines.SaveToFile('c:\body.txt');
  end;


  if not(VarisEmpty(vNextPage)) and not(vNextPage.GetAttribute('disabled',0) = 'True') then
  begin
    FDocumentCompleted := FALSE;
    vNextPage.Click;
    while not FDocumentCompleted do
      Application.HandleMessage;
  end
  else
  begin
    if (aPage > 10) then
    begin
      vNextPage := HTBrowser.OLEObject.Document.GetElementByID(Format('ctl00_CPMain_ucPager_repPages_ctl%.2d_p%d',
            [Ceil(aPage/10),Ceil(aPage/10)]));
      if not(VarisEmpty(vNextPage)) and not(vNextPage.GetAttribute('disabled',0) = 'True') then
      begin
        FDocumentCompleted := FALSE;
        vNextPage.Click;
        while not FDocumentCompleted do
          Application.HandleMessage;

        result := BrowseToNextPage(aPage);
      end
      else
      begin
        result := FALSE;
      end;
    end
    else
    begin
      result := FALSE;
    end;
  end;
end;

function TfrmHattrickKoopjesScanner.ParsePrijzen(var vPlayers:integer):double;
var
  i, j:integer;
  vPrijs, vTempStr:String;
  vTotal,
  vRatio:double;
  vLowestPrize,vHighestPrize,
  v1naLowestPrize,v1naHighestPrize,
  vCurPrize, vAverage: double;
begin
  vPlayers := 0;
  vTotal := 0;
  vLowestPrize := MAXINT;
  v1naLowestPrize := MAXINT;
  v1naHighestPrize := 0;
  vHighestPrize := 0;
  vAverage := 0;
  try
    StreamToMemo(HTBrowser);
    for i:=0 to vMemo.lines.Count -1 do
    begin
      if Pos('€',vMemo.Lines[i]) > 0 then
      begin
        vPrijs := '';
        vTempStr := vMemo.Lines[i];
        for j:=1 to Length(vTempStr) do
        begin
          if (vTempStr[j] in ['0'..'9']) then
          begin
            vPrijs := vPrijs + vTempStr[j];
          end;
        end;
        vCurPrize := StrToFloat(vPrijs);

        if (vCurPrize < vLowestPrize) then
        begin
          v1naLowestPrize := vLowestPrize;
          vLowestPrize := vCurPrize;
        end
        else if (vCurPrize < v1naLowestPrize) then
        begin
          v1naLowestPrize := vCurPrize;
        end;

        if (vCurPrize > vHighestPrize) then
        begin
          v1naHighestPrize := vHighestPrize;
          vHighestPrize := vCurPrize;
        end
        else if (vCurPrize > v1naHighestPrize) then
        begin
          v1naHighestPrize := vCurPrize;
        end;

        vTotal := vTotal + StrToFloat(vPrijs);
        inc(vPlayers);
      end;
    end;
  finally
    vRatio := 0;

    if (vPlayers > 2) then
    begin
      vTotal := vTotal - vLowestPrize - vHighestPrize;
      vPlayers := vPlayers - 2;

      if (vPlayers >= 4) then
      begin
        vRatio := (v1naHighestPrize - v1naLowestPrize) / v1naLowestPrize;
      end;
    end;
    
    if (vPlayers > 0) then
    begin
      vAverage := vTotal / vPlayers;

      //ok.. als er heel erg veel verschil zit tussen minimum en maximum, dan het gemiddelde flink bijstellen
      while (vRatio > 0.50) do
      begin
        if (vRatio < 0.75) then
        begin
          vAverage := ((vAverage * 3) + vLowestPrize) / 4;
        end
        else
        begin
          vAverage := ((vAverage * 2) + vLowestPrize) / 3;
        end;
        vRatio := vRatio - 0.50;
      end;

      if (vPlayers <= 2) then
      begin
        vPlayers := 1;
      end;
    end;

    result := vAverage;
  end;
end;

function TfrmHattrickKoopjesScanner.BerekenMaxPrijs(aTransferprijs: double; aLoon, aLeeftijd, aBlessureWeken: integer): double;
begin
  Result := aTransferprijs * (100 - FTPEMarge) / 100;
  Result := Result * (100 - FMakelaarsPerc) / 100;
  Result := Result * (100 - FVorigeClubs) / 100;

  Result := Result - FKoopjesMarge;
  Result := Result - (aTransferprijs * FKoopjesMargePerc / 100);

  Result := Result - (FKoopjesLoonWeken * aLoon);

  //vanaf 31 jaar per jaar 50k extra winst
  if (aLeeftijd >= 31) then
  begin
    Result := Result - ((aLeeftijd - 30) * 50000);
  end;

  if (aBlessureWeken > 0) then
  begin 
    aLoon := Ceil(aLoon * 1.5) + 15000;
    
    if (aLeeftijd < 27) then
    begin
      //werkelijke tijd = 2/3 van de blessureweken
      Result := Result - (aBlessureWeken * 2 / 3 * aLoon);
    end
    else if (aLeeftijd < 30) then
    begin
      //werkelijke tijd = het aantal blessureweken
      Result := Result - (aBlessureWeken * aLoon);
    end
    else if (aLeeftijd < 32) then
    begin
      //werkelijke tijd = het aantal blessureweken * 1.2
      Result := Result - (aBlessureWeken * 1.3 * aLoon);
    end
    else if (aLeeftijd = 32) then
    begin
      //werkelijke tijd = het aantal blessureweken * 1.4
      Result := Result - (aBlessureWeken * 1.6 * aLoon);
    end
    else if (aLeeftijd = 33) then
    begin
      //werkelijke tijd = het aantal blessureweken * 1.7
      Result := Result - (aBlessureWeken * 2.5 * aLoon);
    end
    else if (aLeeftijd = 34) then
    begin
      //werkelijke tijd = het aantal blessureweken * 3
      if (aBlessureWeken = 1) then //maal 5.5
      begin
        Result := Result - (3 * aLoon);
      end
      else
      begin
        //te veel blessure, gaan we ons niet aan vertillen
        Result := -1;
      end;
    end
    else
    begin
      //te veel blessure, gaan we ons niet aan vertillen
      Result := -1;
    end;
  end;
end;

function TfrmHattrickKoopjesScanner.ParsePlayerInfo:TTSISet;
var
  vBody: TStringList;
begin
  StreamToMemo(HTBrowser);

  vBody := TStringList.Create;
  try
    vBody.Text := vMemo.Text;

    Result := uHattrick.ParsePlayerInfo(vBody);
  finally
    vBody.Free;
  end;
end;

procedure TfrmHattrickKoopjesScanner.SaveKoopje(aScoutingID, aPlayerID, aTPE, aHoogsteBod, aMaxBod: integer; aDeadline: TDateTime;
  vTPEAantalSpelers: integer);
begin
  with uBibDB.CreateInsertSQL(ibdbHTInfo, 'KOOPJES') do
  begin
    try
      ParamByName('ID').AsInteger := uBibDB.GetGeneratorValue(ibdbHTInfo, 'KOOPJES_ID_GEN');
      ParamByName('SCOUTING_ID').AsInteger := aScoutingID;
      ParamByName('PLAYER_ID').AsInteger := aPlayerID;
      ParamByName('TPE').AsInteger := aTPE;
      ParamByName('HOOGSTE_BOD').AsInteger := aHoogsteBod;
      ParamByName('MAX_BOD').AsInteger := aMaxBod;
      ParamByName('DEADLINE').AsDateTime := aDeadline;
      ParamByName('TPE_AANTALSPELERS').AsInteger := vTPEAantalSpelers;
      ParamByName('SCAN_TIJDSTIP').AsDateTime := ESBDates.AddHrs(aDeadline, -1);
      
      ExecQuery;

    finally
      uBibDb.CommitTransaction(Transaction,TRUE);
      Free;
    end;
  end;
end;

function TfrmHattrickKoopjesScanner.SaveScouting(aTSISet:TTSISet; aTalentScouting: boolean): integer;
begin
  if (not aTalentScouting) then
  begin
    uBibDb.ExecSQL(ibdbHTInfo,'DELETE FROM SCOUTING WHERE PLAYER_ID = :ID',
      ['ID'],[aTSISet.PlayerID]);
  end;

  // karakter en rest van kenmerken opslaan
  aTSISet.KarakterID := uHattrick.SaveKarakterProfiel(ibdbHTInfo, aTSISet, False);

  Result := uHattrick.SaveScouting(ibdbHTInfo, aTSISet, aTalentScouting);
end;

end.
