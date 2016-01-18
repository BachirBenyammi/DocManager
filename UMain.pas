unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, ImgList, StdCtrls, OleServer,
  ActnList, XPStyleActnCtrls, ActnMan, Menus, ToolWin, ActnCtrls, CommCtrl,
  Spin, Buttons, ShellApi, ShlObj, FileCtrl, ActiveX, WordXP;

type
  TMainForm = class(TForm)
    PmPrint: TPopupMenu;
    Pageacualle1: TMenuItem;
    Documententier1: TMenuItem;
    pmModels: TPopupMenu;
    ActionManager: TActionManager;
    ActionOpen: TAction;
    ActionNew: TAction;
    ActionModel: TAction;
    ActionSave: TAction;
    ActionSaveAll: TAction;
    ActionSaveAs: TAction;
    ActionPrint: TAction;
    ActionRefresh: TAction;
    ActionClose: TAction;
    ActionCloseAll: TAction;
    ActionFermerWord: TAction;
    ActionHelp: TAction;
    ActionAbout: TAction;
    ActionFirst: TAction;
    ActionLast: TAction;
    ActionNext: TAction;
    ActionEnd: TAction;
    ActionSeek: TAction;
    WordApplication1: TWordApplication;
    SD: TSaveDialog;
    OD: TOpenDialog;
    DocsImageList: TImageList;
    StatusBar: TStatusBar;
    pnlPage: TPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpinEdit1: TSpinEdit;
    pnlDetails: TPanel;
    pnlOpenedDocs: TPanel;
    ToolBarOpenedDocs: TActionToolBar;
    lbOpenedDocs: TListBox;
    Splitter1: TSplitter;
    pnlDocs: TPanel;
    lbDocs: TListBox;
    ToolBarDocs: TActionToolBar;
    ToolBarButtom: TActionToolBar;
    ActionPos: TAction;
    ActionQuit: TAction;
    ActionDel: TAction;
    ActionCopy: TAction;
    ActionMove: TAction;
    ToolsImageList: TImageList;
    pnlExplorer: TPanel;
    PageControl1: TPageControl;
    TabTools: TTabSheet;
    lvTools: TListView;
    Splitter2: TSplitter;
    ActionHide: TAction;
    ActionTop: TAction;
    Splitter3: TSplitter;
    ActionOpenDocFromList: TAction;
    ActionRename: TAction;
    ActionProperties: TAction;
    ActionDocsRefresh: TAction;
    Splitter4: TSplitter;
    pnlAddInfos: TPanel;
    mAddInfos: TMemo;
    ActionEdit: TAction;
    ActionOk: TAction;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure lbDocsClick(Sender: TObject);
    procedure lbDocsDblClick(Sender: TObject);
    procedure lbOpenedDocsClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SpinEdit1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ActionModelExecute(Sender: TObject);
    procedure WordApplication1Quit(Sender: TObject);
    procedure WordApplication1DocumentOpen(ASender: TObject;
      const Doc: _Document);
    procedure WordApplication1NewDocument(ASender: TObject;
      const Doc: _Document);
    procedure WordApplication1DocumentBeforeClose(ASender: TObject;
      const Doc: _Document; var Cancel: WordBool);
    procedure WordApplication1WindowActivate(ASender: TObject;
      const Doc: _Document; const Wn: Window);
    procedure ActionOpenExecute(Sender: TObject);
    procedure ActionNewExecute(Sender: TObject);
    procedure ActionSaveExecute(Sender: TObject);
    procedure ActionSaveAsExecute(Sender: TObject);
    procedure ActionSaveAllExecute(Sender: TObject);
    procedure ActionPrintExecute(Sender: TObject);
    procedure Documententier1Click(Sender: TObject);
    procedure Pageacualle1Click(Sender: TObject);
    procedure ActionRefreshExecute(Sender: TObject);
    procedure ActionCloseExecute(Sender: TObject);
    procedure ActionCloseAllExecute(Sender: TObject);
    procedure ActionFermerWordExecute(Sender: TObject);
    procedure ActionAboutExecute(Sender: TObject);
    procedure ActionFirstExecute(Sender: TObject);
    procedure ActionLastExecute(Sender: TObject);
    procedure ActionNextExecute(Sender: TObject);
    procedure ActionEndExecute(Sender: TObject);
    procedure ActionSeekExecute(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure ActionPosExecute(Sender: TObject);
    procedure ActionQuitExecute(Sender: TObject);
    procedure ActionHideExecute(Sender: TObject);
    procedure lvToolsClick(Sender: TObject);
    procedure ActionTopExecute(Sender: TObject);
    procedure ActionOpenDocFromListExecute(Sender: TObject);
    procedure ActionDelExecute(Sender: TObject);
    procedure ActionMoveExecute(Sender: TObject);
    procedure ActionCopyExecute(Sender: TObject);
    procedure ActionRenameExecute(Sender: TObject);
    procedure ActionPropertiesExecute(Sender: TObject);
    procedure ActionDocsRefreshExecute(Sender: TObject);
    procedure lbOpenedDocsDblClick(Sender: TObject);
  private
    procedure FindFiles(Path, FileName: String; List: TStrings);
    procedure FindDocs(Path, FileName: String; FullList, List : TStrings);
    procedure FindDirs(Path: String; List: TListView);
    procedure FindInfo(FileName: String; List: TStrings);
    procedure ShowActiveProprities(List: TStrings);
    Procedure GetAllOpenedDocs;
    Procedure GetAllModels;
    procedure CloseAllOpenedDocs;
    procedure SaveAllOpenedDocs;
    procedure OpenDocFromList(FileName: string);
    procedure AddDoc;
    procedure NewDoc;
    procedure OpenDoc;
    Procedure SaveDoc;
    Procedure CloseDoc;
    procedure PrintDoc(Range: OleVariant);
    procedure NewDocFromModel(I: Integer);
    procedure SelDoc(I: integer);
    procedure SaveDocAs(I: integer);
    procedure MenuItemClick(Sender: TObject);
    procedure ActivateDoc(I: integer);
    Procedure CloseActiveDoc;
    procedure SaveActiveDoc;
    procedure SaveActiveDocAs;
    procedure PrintActiveDoc (Range: OleVariant);
  end;

var
  MainForm: TMainForm;
  ModelsDir, DocsCurrent: string;
  ModelsList, OpenedDocsList, DocsList: TStrings;
  ActualPage: Integer = 1;
  index: Integer;
  What, Which, Count: OleVariant;
  SaveChanges: OleVariant;
  UnDoc: _Document;
  item: OleVariant;
  
implementation

uses
  USplash;

{$R *.dfm}

function GetMyDocuments: string;
 var
  r: Bool;
  path: array[0..Max_Path] of Char;
 begin
  r := ShGetSpecialFolderPath(0, path, CSIDL_Personal, False) ;
  if not r then 
    raise Exception.Create('Could not find My Documents'' folder location !!') ;
  Result := Path;
 end;

Function DelSlash(Filename:String):String;
begin
  result := FileName;
  if result[length(result)] = '\' then
    result := Copy(result, 0, Length(result) - 1);
end;

function GetParentDir(Filename: String): String;
var
  FirstD, LastD: integer;
begin
  FileName := ExtractFilePath(FileName);
  FirstD := Pos('\', FileName);
  LastD := LastDelimiter('\', FileName);
  if FirstD = LastD then
    result := Copy(FileName, 0 , FirstD - 2)
  else if FirstD < LastD then
    begin
      FileName := DelSlash(FileName);
      LastD := LastDelimiter('\', FileName);
      result := Copy(FileName, LastD + 1, Length(FileName) - LastD + 1);
    end;
end;

procedure Execute (FileName: string);
begin
  ShellExecute(0, 'Open', Pchar(FileName), nil, nil, SW_SHOW);
end;

Function AddSlash(Filename: String): String;
begin
  result := FileName;
  if result[length(result)] <> '\' then
    result := result + '\'
end;

Procedure Exe(Folder: integer);
var
  MyItemIDList: PItemIDList;
  MyShellEx: TShellExecuteInfo;
begin
  SHGetSpecialFolderLocation(0, Folder, MyItemIDList);
  with MyShellEx do
    begin
      cbSize := Sizeof(MyShellEx);
      fMask := SEE_MASK_IDLIST;
      Wnd := 0;
      lpVerb := nil;
      lpFile := nil;
      lpParameters := nil;
      lpDirectory := nil;
      nShow := SW_SHOW;
      hInstApp := 0;
      lpIDList := MyItemIDList
    end;
  ShellExecuteEx(@MyShellEx)
end;

function MyDoc: string;
var
  Path : pchar;
  idList : PItemIDList;
begin
  GetMem(Path, MAX_PATH);
  SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, idList);
  SHGetPathFromIDList(idList, Path);
  Result := AddSlash(string(Path));
  FreeMem(Path);
end;

Function SndFileToRecycleBin(Const FileName: String): Boolean;
var
  Sh: TSHFileOpStructA;
  P1: array[byte] of char;
begin
  FillChar(P1, sizeof(P1), 0);
  StrPcopy(P1, ExpandFileName(FileName) + #0#0);
  with SH do
    begin
      wnd := 0;
      wFunc := FO_DELETE;
      pFrom := P1;
      pTo := nil;
      fFlags := FOF_ALLOWUNDO;
      fAnyOperationsAborted := false;
      hNameMappings := nil
    end;
  Result := (ShFileOperation(Sh) = 0);
end;

Procedure ShowFileProperties(Const filename: String);
Var
  sei: TShellExecuteinfo;
Begin
  FillChar(sei, sizeof(sei), 0);
  sei.cbSize := sizeof(sei);
  sei.lpFile := Pchar(filename);
  sei.lpVerb := 'Properties';
  sei.fMask  := SEE_MASK_INVOKEIDLIST;
  ShellExecuteEx(@sei);
End;

procedure TMainForm.FindFiles(Path, FileName: String; List: TStrings);
var
  FileSR: TSearchRec;
  Result: Integer;
begin
  List.Clear;
  Path := AddSlash(Path);
  Result := FindFirst(Path + FileName, faAnyFile + faHidden + faSysFile +
    faReadOnly, FileSR);
  try
    while Result = 0 do
      begin
        FileName := FileSR.Name;
        FileName := Copy(FileName, 0, Pos(ExtractFileExt(FileName),FileName) - 1);
        if FileName[1] <> '~' then
          List.Add(FileName);
        Result := FindNext(FileSR);
      end;
  finally
    FindClose(FileSR);
  end;
end;

procedure TMainForm.FindDocs(Path, FileName: String; FullList, List: TStrings);
var
  FileSR: TSearchRec;
  Result,i: Integer;
  Ext: TStringList;
begin
  List.Clear;
  FullList.Clear;
  ext := TStringList.Create;
  ext.Add('.doc');
  ext.Add('.xls');
  ext.Add('.ppt');
  Path := AddSlash(Path);
  Result := FindFirst(Path + FileName, faAnyFile + faHidden + faSysFile +
    faReadOnly, FileSR);
  try
    while Result = 0 do
      begin
        FileName := FileSR.Name;
        if (FileName[1] <> '~') and
           ( Ext.IndexOf(ExtractFileExt(FileName)) > -1 ) then
          List.Add(AddSlash(Path) + FileName);
        Result := FindNext(FileSR);
      end;
  finally
    FindClose(FileSR);
  end;
  FullList.AddStrings(List);
  List.Clear;
  for i:= 0 to FullList.count -1 do
    begin
      FileName := ExtractFileName(FullList[i]);
      FileName := Copy(FileName, 0, Pos(ExtractFileExt(FileName),FileName) - 1);
      List.Add(FileName);
    end;
end;

procedure TMainForm.FindDirs(Path: String; List: TListView);
var
  DirSr: TSearchRec;
  Result: Integer;
  NewItem: TListItem;

  function DirNote(Dir: String): Boolean;
  begin
    result := (Dir = '.') or (Dir = '..');
  end;

begin
  List.Clear;
  Path := AddSlash(Path);
  Result := FindFirst(Path  + '*.*', faDirectory, DirSR);
  try
    while Result = 0 do
      begin
        if ((DirSR.Attr and faDirectory) = faDirectory) and not
          DirNote(DirSR.Name) then
          with NewItem do
            begin
              NewItem := List.Items.Add;
              Caption := DirSr.Name;
              ImageIndex := 9;
            end;
        Result := FindNext(DirSR);
      end;
  finally
    FindClose(DirSR);
  end;
end;

procedure TMainForm.FindInfo(FileName: String; List: TStrings);

  function GetTime (ft: _FILETIME) : TDateTime;
  var
    LTime : TFileTime;
    Systemtime : TSystemtime;
  begin
    FileTimeToLocalFileTime( ft, LTime);
    FileTimeToSystemTime( LTime, SystemTime );
    result := SystemTimeToDateTime( SystemTime);
  end;

  function FileSize (f: TSearchRec) : Cardinal;
  begin
    result := f.Size div 1024;
  end;

var
  FileSR: TSearchRec;
  Rslt: Integer;
  
begin
  List.Clear;   
  Rslt := FindFirst(FileName, faAnyFile, FileSR);
  if Rslt = 0 then
  try
    List.add('Name : ' + FileSR.Name);
    List.add('Folder : ' + GetParentDir(FileName));
    List.add('Size : ' + IntToStr(FileSize(FileSR)) + ' KB');
    List.add('Creation date : ' + DateTimeToStr(GetTime(FileSR.FindData.ftCreationTime)));
    List.add('Modification date: ' + DateTimeToStr(GetTime(FileSR.FindData.ftLastWriteTime)));
    List.add('Last access : ' + DateTimeToStr(GetTime(FileSR.FindData.ftLastAccessTime)));
  finally
    FindClose(FileSR);
  end;
end;


procedure TMainForm.FormCreate(Sender: TObject);
begin
  Caption := Application.Title;
  ModelsDir := ExtractFilePath(Application.ExeName);
  ModelsDir := AddSlash(ModelsDir) + 'Models\';

  ModelsList := TStringList.Create;
  OpenedDocsList := TStringList.Create;
  DocsList := TStringList.Create;

  GetAllOpenedDocs;
  GetAllModels;

  PageControl1.TabIndex := 0;
  DocsCurrent := GetMyDocuments;
  FindDocs(DocsCurrent, '*.*', DocsList, lbDocs.Items);
end;

procedure TMainForm.lbDocsClick(Sender: TObject);
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  FindInfo(DocsList[index], mAddInfos.Lines);
end;

procedure TMainForm.OpenDocFromList(FileName: string);
var
  FN: OleVariant;
begin
  with WordApplication1 do
   begin
     Connect;
     Visible := true;
     FN := FileName;
     UnDoc := Documents.Open(FN, EmptyParam, EmptyParam,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
      EmptyParam, EmptyParam);
     UnDoc.Activate;
     Activate;
    end;
end;

procedure TMainForm.lbDocsDblClick(Sender: TObject);
begin
  ActionOpenDocFromListExecute(nil);
end;

procedure TMainForm.lbOpenedDocsClick(Sender: TObject);
begin
  index := lbOpenedDocs.ItemIndex;
  if index = -1 then exit;
  FindInfo(DocsList[index], mAddInfos.Lines);
end;

procedure TMainForm.FormShow(Sender: TObject);
var
  rect : TRect;
begin
 StatusBar.perform(SB_GETRECT, 0, integer(@rect));
 with pnlPage do
  begin
    parent := StatusBar;
    top := rect.top;
    left := rect.left;
    width := rect.right - rect.left;
    height := rect.bottom - rect.top;
    Visible := true;
  end;
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
//  WordApplication1.Disconnect;
//  WordApplication1.Quit;
  ModelsList.Free;
  OpenedDocsList.Free;
  DocsList.Free;
end;

procedure TMainForm.MenuItemClick(Sender: TObject);
begin
  NewDocFromModel((Sender as TMenuItem).Tag);
end;

Procedure TMainForm.GetAllModels;
var
  i: integer;
  NewItem: TMenuItem;
begin
  FindFiles(ModelsDir, '*.dot', ModelsList);
  pmModels.Items.Clear;
  for i:= 0 to ModelsList.Count -1 do
    begin
      NewItem := TMenuItem.Create(pmModels);
      pmModels.Items.Add(NewItem);
      with NewItem do
        begin
          Caption := ModelsList[i];
          OnClick := MenuItemClick;
          Tag := i;
        end;
    end;
end;

Procedure TMainForm.GetAllOpenedDocs;
var
  I: integer;
  index: OleVariant;
begin
  WordApplication1.Connect;
  OpenedDocsList.Clear;
  lbOpenedDocs.Clear;
  index := -1;
  with WordApplication1.Documents do
    for I:= 1 to Count do
      begin
        index := I;
        UnDoc := item(index);
        AddDoc;
      end;
end;

procedure TMainForm.SelDoc(I: integer);
begin
  SendMessage(lbOpenedDocs.Handle, messages.LB_SETCURSEL ,I, 0);
end;

procedure TMainForm.CloseAllOpenedDocs;
var
  I: OLEVariant;
begin
  I:= 1;
  with WordApplication1.Documents do
    while Count >= 1 do
      begin
        UnDoc := Item(I);
        CloseDoc;
      end;
  OpenedDocsList.Clear;
  lbOpenedDocs.Clear;
end;

procedure TMainForm.SaveAllOpenedDocs;
var
  I: integer;
  index: OleVariant;
begin
  with WordApplication1.Documents do
    for I:= 1 to Count  do
      begin
        index := I;
        UnDoc := item(index);
        SaveDoc;
      end;
end;

procedure TMainForm.OpenDoc;
var
    FileName: OleVariant;
begin
  OD.InitialDir := DocsCurrent;
  if OD.Execute then
    with WordApplication1 do
     begin
       Connect;
       FileName := OD.FileName;
       Visible := true;
       UnDoc := Documents.Open(FileName, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam);
       UnDoc.Activate;
       Activate;
      end;
end;

function ExtractFileNameWE (FileName:string): string;
var
  I: integer;
begin
  I := Pos(ExtractFileExt(FileName),FileName) - 1;
  if I <> -1 then
    FileName := Copy(FileName, 0, I);
  Result := FileName;
end;

procedure TMainForm.AddDoc;
var
  FileName: string;
begin
  FileName := UnDoc.Name;
  OpenedDocsList.Add(UnDoc.FullName);
  lbOpenedDocs.Items.Add(ExtractFileNameWE(FileName));
  Caption := 'Doc Manager [' + FileName + ']';
  index := lbOpenedDocs.Count -1;
  if index = -1 then exit;
  SelDoc(Index);
end;

procedure TMainForm.NewDoc;
begin
  with WordApplication1 do
    begin
      Connect;
      Visible := True;
      UnDoc := Documents.Add(EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      UnDoc.Activate;
      Activate;
    end;
end;

procedure TMainForm.NewDocFromModel(I: Integer);
var
  Template, AsModel: OleVariant;
begin
  With WordApplication1 do
    begin
      Connect;
      Visible := True;
      Template := ModelsDir + ModelsList[I] + '.dot';
      AsModel := False;
      UnDoc := Documents.Add(Template, AsModel, EmptyParam, EmptyParam);
      UnDoc.Activate;
      Activate;
    end;
end;

Procedure TMainForm.CloseActiveDoc;
begin
  UnDoc := WordApplication1.ActiveDocument;
  CloseDoc;
end;

Procedure TMainForm.CloseDoc;
var
  Confirm: boolean;
  OriginalFormat, RouteDocument: OleVariant;
begin
  OriginalFormat := UnAssigned;
  RouteDocument := UnAssigned;
  Confirm := True;
  with UnDoc do
    begin
      if not Saved then
        case MessageDlg('Save changes to ' + Name,
          mtConfirmation, [MbYes, MbNo, MbCancel], 0) of
          MrYes : SaveChanges := wdSaveChanges;
          MrNo :  SaveChanges := wdDoNotSaveChanges;
          MrCancel : Confirm := false;
        end;
      if Confirm then
        begin
          Close(SaveChanges, OriginalFormat, RouteDocument);
          Caption := 'Doc Manager';
        end;
    end;
end;

procedure TMainForm.SaveDocAs(I: integer);
var
  FileName : OLEVariant;
begin
  SD.InitialDir := DocsCurrent;
  item := ExtractFileName(OpenedDocsList[i]);
  if SD.Execute then
    begin
      FileName := SD.FileName;
      WordApplication1.Documents.Item(item).SaveAs(FileName, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam);
      GetAllOpenedDocs;        
    end;
end;

procedure TMainForm.SaveDoc;
begin
  if UnDoc.FullName = UnDoc.Name then
    SaveDocAs(lbOpenedDocs.Items.IndexOf(ExtractFileNameWE(UnDoc.Name)) + 1)
  else
    UnDoc.Save;
end;

procedure TMainForm.SaveActiveDoc;
begin
  UnDoc := WordApplication1.ActiveDocument;
  SaveDoc;
end;

procedure TMainForm.SaveActiveDocAs;
var
  FileName : OLEVariant;
begin
  SD.InitialDir := DocsCurrent;
  UnDoc := WordApplication1.ActiveDocument;
  if SD.Execute then
    begin
      FileName := SD.FileName;
      UnDoc.SaveAs(FileName, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    end;
end;

procedure TMainForm.ActivateDoc(I: integer);
begin
  item := ExtractFileName(OpenedDocsList[i]);
  with WordApplication1 do
    begin
      UnDoc := Documents.Item(item);
      UnDoc.Activate;
      Activate;
    end;
  FindInfo(OpenedDocsList[i], mAddInfos.Lines);
  ShowActiveProprities(mAddInfos.Lines);
end;

procedure TMainForm.ShowActiveProprities(List: TStrings);
begin
end;

procedure TMainForm.PrintDoc(Range: OleVariant);
var
  Background: OleVariant;
  PageType: OleVariant;
begin
  Background := False;
  PageType := wdPrintAllPages;
  UnDoc.PrintOut(Background, EmptyParam, Range,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
    PageType, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam);
end;

procedure TMainForm.PrintActiveDoc (Range: OleVariant);
begin
  UnDoc := WordApplication1.ActiveDocument;
  PrintDoc(Range);
end;

procedure TMainForm.SpinEdit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = Vk_Return then
    ActionSeekExecute(nil);
end;

procedure TMainForm.ActionModelExecute(Sender: TObject);
begin
  GetAllModels;
  With Mouse.CursorPos do
    pmModels.Popup(X , Y);
end;

procedure TMainForm.WordApplication1Quit(Sender: TObject);
begin
  WordApplication1.Disconnect;
end;

procedure TMainForm.WordApplication1DocumentOpen(ASender: TObject;
  const Doc: _Document);
begin
  GetAllOpenedDocs;
end;

procedure TMainForm.WordApplication1NewDocument(ASender: TObject;
  const Doc: _Document);
begin
  GetAllOpenedDocs;
end;

procedure TMainForm.WordApplication1DocumentBeforeClose(ASender: TObject;
  const Doc: _Document; var Cancel: WordBool);
begin
  index := lbOpenedDocs.Items.IndexOf(ExtractFileNameWE(Doc.Name));
  if index = -1 then exit;
  OpenedDocsList.Delete(index);
  lbOpenedDocs.Items.Delete(index);
end;

procedure TMainForm.WordApplication1WindowActivate(ASender: TObject;
  const Doc: _Document; const Wn: Window);
begin
  index := lbOpenedDocs.Items.IndexOf(ExtractFileNameWE(Doc.Name));
  if index = -1 then exit;
  SelDoc(index);
end;

procedure TMainForm.ActionOpenExecute(Sender: TObject);
begin
  OpenDoc;
end;

procedure TMainForm.ActionNewExecute(Sender: TObject);
begin
  NewDoc;
end;

procedure TMainForm.ActionSaveExecute(Sender: TObject);
begin
  SaveActiveDoc;
end;

procedure TMainForm.ActionSaveAsExecute(Sender: TObject);
begin
  SaveActiveDocAs;
end;

procedure TMainForm.ActionSaveAllExecute(Sender: TObject);
begin
  SaveAllOpenedDocs;
end;

procedure TMainForm.ActionPrintExecute(Sender: TObject);
begin
  With Mouse.CursorPos do
    pmPrint.Popup(X , Y);
end;

procedure TMainForm.Documententier1Click(Sender: TObject);
begin
  PrintActiveDoc (wdPrintAllDocument);
end;

procedure TMainForm.Pageacualle1Click(Sender: TObject);
begin
  PrintActiveDoc (wdPrintCurrentPage);
end;

procedure TMainForm.ActionRefreshExecute(Sender: TObject);
begin
 GetAllOpenedDocs;
end;

procedure TMainForm.ActionCloseExecute(Sender: TObject);
begin
  CloseActiveDoc;
  GetAllOpenedDocs;
end;

procedure TMainForm.ActionCloseAllExecute(Sender: TObject);
begin
  CloseAllOpenedDocs;
end;

procedure TMainForm.ActionFermerWordExecute(Sender: TObject);
begin
  if WordApplication1.Documents.Count > 0  then
    begin
      case MessageDlg('Quit Microsoft Word,' +
        ' Do you want to save changes ?', mtConfirmation,
        [MbYes, MbNo, MbCancel], 0) of
        MrYes : SaveChanges := wdSaveChanges;
        MrNo :  SaveChanges := wdDoNotSaveChanges;
        MrCancel :Exit;
      end;
      WordApplication1.Quit(SaveChanges);
      OpenedDocsList.Clear;
      lbOpenedDocs.Clear;
    end;
end;

procedure TMainForm.ActionAboutExecute(Sender: TObject);
begin
  SplashForm := TSplashForm.Create(Application);
  try
    SplashForm.BtnClose.Visible := true;;
    SplashForm.ShowModal;
  finally
    SplashForm.Free
  end;
end;

procedure TMainForm.ActionFirstExecute(Sender: TObject);
begin
  What := WdGotoPage;
  Which := wdGoToFirst;
  WordApplication1.Selection.GoTo_(What, Which, Count, EmptyParam);
  SpinEdit1.Value := Count;
end;

procedure TMainForm.ActionLastExecute(Sender: TObject);
begin
  WordApplication1.Selection.GoToPrevious(WdGotoPage);
  SpinEdit1.Value := WdGotoPage;
end;

procedure TMainForm.ActionNextExecute(Sender: TObject);
begin
  WordApplication1.Selection.GoToNext(WdGotoPage);
  SpinEdit1.Value := WdGotoPage;
end;

procedure TMainForm.ActionEndExecute(Sender: TObject);
begin
  What := WdGotoPage;
  Which := wdGoToLast;
  WordApplication1.Selection.GoTo_(What, Which, Count, EmptyParam);
  SpinEdit1.Value := Count;
end;

procedure TMainForm.ActionSeekExecute(Sender: TObject);
begin
  What := WdGotoPage;
  Which := wdGotoAbsolute;
  Count := spinEdit1.Value;
  WordApplication1.Selection.GoTo_(What, Which, Count, EmptyParam);
  SpinEdit1.Value := Count;
end;

procedure TMainForm.FormResize(Sender: TObject);
var
  rect : TRect;
begin
 StatusBar.perform(SB_GETRECT, 1, integer(@rect));
 with ToolBarButtom do
  begin
    Visible := false;
    parent := StatusBar;
    top := rect.top;
    left := rect.Right  - width;
    height := rect.bottom - rect.top;
    Visible := true;
  end;
end;

procedure TMainForm.ActionPosExecute(Sender: TObject);
begin
  case Align of
    alNone:
      begin
        Align := alRight;
        ActionPos.Caption := 'Normal';
        ActionPos.Checked := True;
        ActionPos.Hint := 'Restore Window';
        SetWindowPos(FindWindow('OpusApp',nil), HWND_TOP , 0, 0,
          Screen.WorkAreaWidth - Width, Height, SWP_SHOWWINDOW);
        end;
    alRight:
      begin
        Align := alNone;
        SetBounds((Screen.DesktopWidth - Width) div 2,
          (Screen.DesktopHeight - Height) div 2, 450, 500);
        ActionPos.Caption := 'Left';
        ActionPos.Checked := False;
        ActionPos.Hint := 'Extend application';
      end;
  end;
end;

procedure TMainForm.ActionQuitExecute(Sender: TObject);
begin
  if MessageDlg('Do you want to quit ' + Application.Title + ' ?', mtConfirmation,
    [MbYes, MbCancel], 0) = mrYes then
      Application.Terminate;
end;

procedure TMainForm.ActionHideExecute(Sender: TObject);
begin
  case pnlExplorer.Visible of
    true:
      begin
        pnlExplorer.Visible := false;
        ActionHide.Caption := 'Show';
        ActionHide.ImageIndex := 15;
      end;
    false:
      begin
        pnlExplorer.Visible := true;
        pnlExplorer.Align := alRight;
        ActionHide.Caption := 'Hide';
        ActionHide.ImageIndex := 16;
      end
  end;
end;

procedure TMainForm.lvToolsClick(Sender: TObject);
begin
  case lvTools.ItemIndex of
    0: Exe(CSIDL_PERSONAL);
    1: Execute('winword.exe');
  end;
end;

procedure TMainForm.ActionTopExecute(Sender: TObject);
begin
  case FormStyle of
    fsNormal:
      begin
        FormStyle := fsStayOnTop;
        ActionTop.Caption := 'On top';
        ActionTop.ImageIndex := 38;
      end;
    fsStayOnTop:
      begin
        FormStyle := fsNormal;
        ActionTop.Caption := 'Not on top';
        ActionTop.ImageIndex := 39;
      end
  end;
end;

procedure TMainForm.ActionOpenDocFromListExecute(Sender: TObject);
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  OpenDocFromList(DocsList[index]);
end;

procedure TMainForm.ActionDelExecute(Sender: TObject);
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  SndFileToRecycleBin(DocsList[index]);
  FindDocs(DocsCurrent, '*.*', DocsList, lbDocs.Items);
end;

procedure TMainForm.ActionMoveExecute(Sender: TObject);
var
  Folder: string;
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  Folder := DocsCurrent;
  if SelectDirectory('Select a folder', '', Folder) then
    begin
      MoveFile(Pchar(DocsList[index]),
        Pchar(AddSlash(Folder) + ExtractFileName(DocsList[index])));
      FindDocs(DocsCurrent, '*.*', DocsList, lbDocs.Items);
    end;
end;

procedure TMainForm.ActionCopyExecute(Sender: TObject);
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  SD.InitialDir := DocsCurrent;
  SD.FileName := DocsList[index];
  if SD.Execute then
      CopyFile(Pchar(DocsList[index]),
        pchar(SD.FileName), False);
end;

procedure TMainForm.ActionRenameExecute(Sender: TObject);
var
  OldFileName, FileName: string;
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  OldFileName := ExtractFileName(DocsList[index]);
  FileName := OldFileName;
  if InputQuery('Rename document', 'Please give a name', FileName) then
    begin
      if Pos(ExtractFileExt(FileName),FileName) <= 0 then
        FileName := FileName + '.doc';
      RenameFile(DocsCurrent + OldFileName, DocsCurrent + FileName);
      FindDocs(DocsCurrent, '*.*', DocsList, lbDocs.Items);
    end;
end;

procedure TMainForm.ActionPropertiesExecute(Sender: TObject);
begin
  index := lbDocs.ItemIndex;
  if index = -1 then exit;
  ShowFileProperties(DocsList[index]);
end;

procedure TMainForm.ActionDocsRefreshExecute(Sender: TObject);
begin
  FindDocs(DocsCurrent, '*.*', DocsList, lbDocs.Items);
end;

procedure TMainForm.lbOpenedDocsDblClick(Sender: TObject);
begin
 ActivateDoc(lbOpenedDocs.ItemIndex);
end;

end.
