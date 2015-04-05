unit UMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, WordXP, Vcl.OleServer, Clipbrd,
  Vcl.ExtCtrls, pngimage, FileCtrl, Vcl.Samples.Spin, OfficeXP, Math;

type
  TFMain = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    edSelectPictures: TEdit;
    edSelectDoc: TEdit;
    btnSelectPictures: TButton;
    btnSelectDoc: TButton;
    WordApplication: TWordApplication;
    WordDocument: TWordDocument;
    btnGenerate: TButton;
    btnExit: TButton;
    btnCount: TButton;
    lblStatus: TLabel;
    PictureImage: TImage;
    SaveDialog: TSaveDialog;
    GroupBox1: TGroupBox;
    seLeft: TSpinEdit;
    Label3: TLabel;
    Label4: TLabel;
    seRight: TSpinEdit;
    Label5: TLabel;
    seTop: TSpinEdit;
    Label6: TLabel;
    seBottom: TSpinEdit;
    Label7: TLabel;
    seScale: TSpinEdit;
    lblWord: TLabel;
    procedure btnExitClick(Sender: TObject);
    procedure btnGenerateClick(Sender: TObject);
    procedure btnCountClick(Sender: TObject);

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSelectDocClick(Sender: TObject);
    procedure btnSelectPicturesClick(Sender: TObject);
  private
    { Private declarations }
    FilesList, StationsList: TStringList;
    Counted: boolean;
    function GetPathToPictures: string;

    property PathToPictures: string read GetPathToPictures;
    procedure CountFiles;

    function CentimetersToPoints(aValue: Single): Single;
    function MillimetersToPoints(aValue: Single): Single;
    function PixelsToPoints(aValue: Single): Single;
  public
    { Public declarations }
  end;

var
  FMain: TFMain;

implementation

{$R *.dfm}

procedure TFMain.btnCountClick(Sender: TObject);
begin
  CountFiles;
end;

procedure TFMain.btnExitClick(Sender: TObject);
begin
  Close;
end;

function TFMain.CentimetersToPoints(aValue: Single): Single;
begin
  Result := 28.35 * aValue;
end;

procedure TFMain.btnGenerateClick(Sender: TObject);
var
  i, j, FilesPerStation: integer;
  Shapes: InlineShapes;
  Range: WordRange;
begin
  if not Counted then
    CountFiles;
  if StationsList.Count = 0 then
    exit;
  if FilesList.Count mod StationsList.Count <> 0 then
    if MessageBox(0, 'Number of files doesn''t correspond number of station. Continue?',
         'Confirm', MB_ICONWARNING or MB_YESNO or MB_DEFBUTTON1) = mrNo then
      exit;

  PictureImage.Visible := false;
  lblWord.Visible := true;
  WordApplication.Connect;
  try
    try
      WordApplication.Options.CheckSpellingAsYouType := false;
      WordApplication.Options.CheckGrammarAsYouType := false;

      WordApplication.Documents.Add(EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      WordDocument.ConnectTo(WordApplication.ActiveDocument);
      try
        WordDocument.PageSetup.TopMargin := MillimetersToPoints(seTop.Value);
        WordDocument.PageSetup.BottomMargin := MillimetersToPoints(seBottom.Value);
        WordDocument.PageSetup.LeftMargin := MillimetersToPoints(seLeft.Value);
        WordDocument.PageSetup.RightMargin := MillimetersToPoints(seRight.Value);

        WordDocument.Tables.Add(WordDocument.Range, FilesList.Count + StationsList.Count, 1, EmptyParam, EmptyParam);

        lblWord.Visible := false;
        PictureImage.Visible := true;
        FilesPerStation := FilesList.Count div StationsList.Count;
        for i := 0 to StationsList.Count - 1 do
        begin
          Range := WordDocument.Tables.Item(1).Cell((i * (FilesPerStation + 1)) + 1, 1).Range;
          Range.Text := StationsList[i];
          Range.ParagraphFormat.Alignment := wdAlignParagraphCenter;

          for j := 0 to FilesPerStation - 1 do
          begin
            Range := WordDocument.Tables.Item(1).Cell((i * (FilesPerStation + 1)) + 2 + j, 1).Range;
            Range.Select;
            Shapes := WordApplication.Selection.InlineShapes;

            Shapes.AddPicture(PathToPictures + FilesList[(i * FilesPerStation) + j], false, true, Range);
            WordApplication.Selection.GoToPrevious(wdGoToGraphic);
            WordApplication.Selection.Expand(wdCharacter);
            Shapes.Item(Shapes.Count).LockAspectRatio := msoTrue;

            PictureImage.Picture.LoadFromFile(PathToPictures + FilesList[(i * FilesPerStation) + j]);

            Shapes.Item(Shapes.Count).Width :=
              Ceil(PixelsToPoints(PictureImage.Picture.Width * seScale.Value / 100));
            Shapes.Item(Shapes.Count).Height :=
              Ceil(PixelsToPoints(PictureImage.Picture.Height * seScale.Value / 100));

            Application.ProcessMessages;
          end;
        end;

        WordDocument.Tables.Item(1).Rows.Item(1).Select;
        for i := 1 to StationsList.Count do
        begin
          try
            WordApplication.Selection.MoveDown(wdLine, i * (FilesPerStation + 1), wdMove);
          except on E: Exception do
            WordApplication.Selection.MoveDown(wdLine, i * (FilesPerStation), wdMove);
          end;
          WordApplication.Selection.InsertBreak(wdPageBreak);
          Application.ProcessMessages;
        end;

        WordDocument.SaveAs(edSelectDoc.Text);
        WordApplication.Visible := true;
      finally
        WordDocument.Disconnect;
      end;
    except
      on E: Exception do
      begin
        WordApplication.Quit;
        raise;
      end;
    end;
  finally
    WordApplication.Disconnect;
  end;
end;

procedure TFMain.btnSelectDocClick(Sender: TObject);
begin
  SaveDialog.InitialDir := ExtractFilePath(edSelectDoc.Text);
  SaveDialog.FileName := edSelectDoc.Text;
  if SaveDialog.Execute then
    edSelectDoc.Text := SaveDialog.FileName;
end;

procedure TFMain.btnSelectPicturesClick(Sender: TObject);
var
  InputDir: string;
begin
  InputDir := edSelectPictures.Text;
  if SelectDirectory('Input dir select', '', InputDir, []) then
  begin
    edSelectPictures.Text := InputDir;
    Counted := false;
  end;
end;

procedure TFMain.FormCreate(Sender: TObject);
begin
  FilesList := TStringList.Create;
  StationsList := TStringList.Create;
  Counted := false;
  edSelectPictures.Text := ExtractFilePath(Application.ExeName);
  edSelectDoc.Text := ExtractFilePath(Application.ExeName) + 'test.doc';
end;

procedure TFMain.FormDestroy(Sender: TObject);
begin
  FilesList.Free;
  StationsList.Free;
end;

function TFMain.GetPathToPictures: string;
begin
  Result := IncludeTrailingPathDelimiter(edSelectPictures.Text);
end;

function TFMain.MillimetersToPoints(aValue: Single): Single;
begin
  Result := CentimetersToPoints(aValue / 10);
end;

function TFMain.PixelsToPoints(aValue: Single): Single;
begin
  Result := aValue * 72 / 96;
end;

procedure TFMain.CountFiles;
var
  SR: TSearchRec;
  Station: string;
begin
  FilesList.Clear;
  StationsList.Clear;
  if FindFirst(PathToPictures + '*.*', faAnyFile and not faDirectory, SR) = 0 then
    try
      repeat
        Station := ExtractFileName(SR.Name);
        FilesList.Add(Station);
        Station := Copy(Station, 1, 4);
        if StationsList.IndexOf(Station) = -1 then
          StationsList.Add(Station);
      until FindNext(SR) <> 0;
    finally
      FindClose(SR);
    end;
  lblStatus.Caption := Format('%d files for %d stations were found in directory', [FilesList.Count, StationsList.Count]);
  Counted := true;
end;

end.
