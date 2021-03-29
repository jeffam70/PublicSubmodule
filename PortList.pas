unit PortList;

{--------------------------------------------------------------------------------------------------------------------------------+
ฆ                                                        COPYRIGHT NOTICE                                                        ฆ
ฆ--------------------------------------------------------------------------------------------------------------------------------ฆ
ฆ UNIT      PortList.pas                                                                                                         ฆ
ฆ AUTHOR    Jeff Martin                                                                                                          ฆ
ฆ COPYRIGHT (c) 2021 Parallax Inc.                                                                                               ฆ
ฆ--------------------------------------------------------------------------------------------------------------------------------ฆ
ฆ                                        PERMISSION NOTICE (TERMS OF USE): MIT X11 License                                       ฆ
ฆ--------------------------------------------------------------------------------------------------------------------------------ฆ
ฆ Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation     ฆ
ฆ files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy,     ฆ
ฆ modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software ฆ
ฆ is furnished to do so, subject to the following conditions:                                                                    ฆ
ฆ                                                                                                                                ฆ
ฆ The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. ฆ
ฆ                                                                                                                                ฆ
ฆ THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE           ฆ
ฆ WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR          ฆ
ฆ COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,    ฆ
ฆ ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                          ฆ
+--------------------------------------------------------------------------------------------------------------------------------}

{This is the Parallax PortList object.  It handles everything regarding enumeration of serial ports and includes features to allow the application and
 user to sort ports and filter ports out of the logical "search list" either by Port ID or Port Description rules.

 This class creates a form called "PortListForm" (allowing the user to manage port search rules) and a vital object called "COM" (providing the
 interface to the serial port list).

 To use the Port List class:
   1)  Add "PortList" to the uses clause in the appropriate unit(s).
   2)  Set "PortListForm.DeviceName" to the proper device name (shows up on Serial Port Search List form);
   3)  Set "COM.OnReadPortRules" to a read-port-rules event handler.  This event occurs when the PortList class needs to read the external port list
       rules preference string.  Make sure to write the read-port-rules event handler to return the default port rules if the Default parameter is True.
       See TReadPortRulesEvent declaration for read-port-rules' function syntax.
       NOTE: Setting this event property causes the event to trigger immediately to retrieve the initial value.
   4)  Set "COM.OnWritePortRules" to a write-port-rules event hander.  This event occurs when the PortList class needs to write the external port list
       rules preference string.
       See TWritePortRulesEvent declaration for write-port-rules' procedure syntax.
   5)  Call "PortListForm.ShowModal" to display the Serial Port Search List form to allow user to configure desired preferences.  The user's actions may
       cause TReadPortRulesEvent and TWritePortRulesEvent events.
       NOTE: The method above is recommended for most cases where port list rules need to be changed, however, to change them externally, modify the
       COM.RuleString property, or the individual strings in COM.Rules.
   6)  Set "COM.Filtered" property to True to have the COM object filter the list according to preferences.
   7)  Set "COM.Scannable" property to True to have the COM object filter the list to contain only scannable ports (rather than ports that are present but
       excluded by preferences, or ports that are not present, but included by rules).  NOTE: the Filtered property must be true for the Scannable property
       to have an effect.
   8)  Call "COM.Refresh" to have the COM object refresh its list of available serial ports.
   9)  Reference "COM.Count", "COM.PortID[]", "COM.PortsExcluded", "COM.IndexOfPortID()", "COM.PortDesc[]", and other properties as needed to reference the
       list of available ports.
       Remember: Always call COM.Refresh before starting to reference these properties to ensure they are up-to-date.
   10) To be notified when a serial port add/remove event has occurred, set "COM.OnDeviceChange" to an appropriate event handler.
       See the TDeviceChangeEvent declaration for device-change procedure syntax.
}

interface

uses
  Windows, Messages, SysUtils, StrUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, Grids, Math, Menus, StdCtrls, ExtCtrls,
  SetupApi, MultiMon;

type
  {Forward declarations}
  TPortDescForm = class;
  TPHintWindow = class;

  {WM_DEVICECHANGE message structures for detecting serial port changes (USB port add/remove)}
  PDEV_BROADCAST_HDR = ^DEV_BROADCAST_HDR;
  DEV_BROADCAST_HDR = packed record
    dbch_size       : DWORD;
    dbch_devicetype : DWORD;
    dbch_reserved   : DWORD;
  end;

  PDEV_BROADCAST_PORT = ^DEV_BROADCAST_PORT;
  DEV_BROADCAST_PORT = packed record
    dbcp_size       : DWORD;
    dbcp_devicetype : DWORD;
    dbcp_reserved   : DWORD;
    dbcp_name       : char;
  end;

  TDevMsg = packed record
    WinHandle : HWND;
    Event     : Cardinal;
    case integer of
      0: (Header : PDEV_BROADCAST_HDR);
      1: (Port   : PDEV_BROADCAST_PORT);
  end;


  {Port List Form}
  TPortListForm = class(TForm)
    CustomPortList: TStringGrid;
    PortsListPopup: TPopupMenu;
    ExcludePortItem: TMenuItem;
    ExcludeAllPortsByDefaultItem: TMenuItem;
    IncludePortItem: TMenuItem;
    FilterPortsByDescriptionItem: TMenuItem;
    N1: TMenuItem;
    IncludeAllPortsByDefaultItem: TMenuItem;
    CancelButton: TButton;
    HeaderLabel: TLabel;
    InstructionPanel: TPanel;
    Label2: TLabel;
    Shape1: TShape;
    Shape2: TShape;
    ItalicHeaderLabel: TLabel;
    ClickAndDragLabel: TLabel;
    RightClickLabel: TLabel;
    AcceptButton: TButton;
    N2: TMenuItem;
    ClearAllRulesItem: TMenuItem;
    UndoLastChangeItem: TMenuItem;
    RedoPriorChangeItem: TMenuItem;
    RemovePortRuleItem: TMenuItem;
    RemovePortDescriptionRuleItem: TMenuItem;
    EditPortDescriptionRuleItem: TMenuItem;
    RestoreDefaultsButton: TButton;
    procedure WMDeviceChange(var Msg: TDevMsg); message WM_DEVICECHANGE;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AcceptButtonClick(Sender: TObject);
    procedure RestoreDefaultsButtonClick(Sender: TObject);
    procedure CustomPortListDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure CustomPortListMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure CustomPortListMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure CustomPortListMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure CustomPortListRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    function  CustomPortListCanHint(Sender: TObject; ItemID: Integer; var Text: String): Boolean;
    procedure CustomPortListContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure PopupMenuItemClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FDeviceName   : String;
    FPortDescForm : TPortDescForm;                                {Child form (Port Description edit form)}
    FCPLHint      : TPHintWindow;                                 {The custom hint window that appears for mouse-over-port-items}
    function  GetDisplayableWindowPosition(Bounds: TRect): TRect;
    procedure EnsureWindowDisplayable(Window: TForm);
    procedure SizeComponents;
    procedure UpdateDisplay(NoRefresh: Boolean = False);
    procedure Undo(PerformUndo: Boolean);
    procedure Redo;
    procedure FinalizeModifications;
  public
    { Public declarations }
    property DeviceName : String read FDeviceName write FDeviceName;
  end;


  {Port List Description Form}
  TPortDescForm = class(TCustomForm)
    PurposeLabel: TLabel;
    FilterComboBox: TComboBox;
    DescriptionLabel: TLabel;
    InstructionPanel: TPanel;
    TipsLabel: TLabel;
    Bullet1: TShape;
    Bullet2: TShape;
    Bullet3: TShape;
    Bullet4: TShape;
    GenSpecLabel: TLabel;
    AsteriskLabel: TLabel;
    DescCaseLabel: TLabel;
    OverrideLabel: TLabel;
    DescriptionEdit: TEdit;
    MatchesLabel: TLabel;
    MatchCountLabel: TLabel;
    OkayButton: TButton;
    CancelButton: TButton;
    procedure FormActivate(Sender: TObject);
    procedure FilterComboBoxChange(Sender: TObject);
    procedure DescriptionEditKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OkayButtonClick(Sender: TObject);
  private
    { Private declarations }
    MatchCount : Integer;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    function ShowModal(RuleIdx: Integer; Description: String): Integer;  reintroduce;
    procedure UpdatePurpose;
    procedure UpdateMatchCount;
  end;


  {Port State enumeration}
  TPortState = (psAbsentIncluded, psAbsentExcluded, psPresentIncluded, psPresentExcluded);
  TPortEvent = (peAdded, peRemoved);

  {Port Properties Structure}
  PPortProperties = ^TPortProperties;
  TPortProperties = record
    State       : TPortState;                                 {Indicates the present/absent and included/excluded state of the port}
    InclExcl    : Boolean;                                    {True = included or excluded by a rule, False = no rule yet applied}
    InclExclIdx : Integer;                                    {If InclExcl = True, indicates the index of rule that was applied}
  end;

  {Port Metrics Object}
  TDeviceChangeEvent = procedure (Sender: TObject; PortEvent: TPortEvent; PortID: String) of object;
  TReadPortRulesEvent = function (Sender: TObject; Default: Boolean): String of object;
  TWritePortRulesEvent = procedure (Sender: TObject; PortRules: String) of object;

  TPortMetrics = class(TObject)
    FOSPorts          : TStrings;                               {Pure list of ports (entire description including Port ID) as reported by O.S.}
    FPorts            : TStrings;                               {Filtered list of ports (entire description including Port ID)}
    FPortIDs          : TStrings;                               {Filtered list of port IDs}
    FPortDescs        : TStrings;                               {Filtered list of port descriptions (without Port ID)}
    FProperties       : TList;                                  {Filtered list of port Properties}
    FRules            : TStrings;                               {List of rules guiding this set of ports}
    FFiltered         : Boolean;                                {False = show pure list of ports as reported by 0.S., True = show filtered ports (rules applied)}
    FScannable        : Boolean;                                {False = show all ports, True = show scannable ports only (those allowed by rules)}
    FPortsExcluded    : Integer;                                {Indicates number of "present" ports excluded via rules}
    FOnDeviceChange   : TDeviceChangeEvent;                     {Event that triggers when a serial port is added or removed from the system}
    FOnReadPortRules  : TReadPortRulesEvent;                    {Event that triggers when TPortMetrics needs to read external Port Rules preference}
    FOnWritePortRules : TWritePortRulesEvent;                   {Event that triggers when TPortMetrics needs to write external Port Rules preference}
  private
    procedure DeviceChanged(PortEvent: TPortEvent; PortID: String);
    procedure DoReadPortRules(Default: Boolean = False);
    procedure DoWritePortRules;
    procedure EnumerateComPorts(LongNames: Boolean; var Ports: TStrings);
    function  ExtractPortID(Port: String): String;
    function  ExtractPortDescription(Port: String): String;
    procedure FilterPorts;
    procedure ParseSearchRuleString(RuleStr: String);
    function  GenerateSearchRuleString: String;
    procedure RulesChanged(Sender: TObject);
    function  GetCount: Integer;
    function  GetGlobal(Index: Integer): Boolean;
    function  GetPort(Index, DataType: Integer): String;
    function  GetProperty(Index, DataType: Integer): Boolean;
    function  GetDescRulesString: String;
    function  GetIDRulesString: String;
    function  GetSortRulesString: String;
    function  GetRuleIdx(Index: Integer): Integer;
    function  GetRuleType(Index, DataType: Integer): Boolean;
    procedure SetFiltered(Value: Boolean);
    procedure SetScannable(Value: Boolean);
    procedure SetFOnReadPortRules(Value: TReadPortRulesEvent);
  public
    constructor Create;
    destructor  Destroy;  reintroduce;
    property Count : Integer read GetCount;                                                           {Count of FPorts/FPortIDs/FPortDescs (affected by state of FFiltered and FScannable)}
    property Excluded[Index: Integer] : Boolean index 2 read GetProperty;                             {True = excluded, False = not excluded}
    property Filtered : Boolean read FFiltered write SetFiltered;                                     {True = FPorts/FPortIDs/FPortDescs contains data filtered by rules, False = contains pure "present" port list}
    property GlobalExclude : Boolean index 1 read GetGlobal;                                          {True = Ports excluded by default, False = Ports included by default}
    property GlobalInclude : Boolean index 0 read GetGlobal;                                          {True = Ports included by default, False = Ports excluded by default}
    property OnDeviceChange : TDeviceChangeEvent read FOnDeviceChange write FOnDeviceChange;          {Event for device change notification}
    property OnReadPortRules : TReadPortRulesEvent read FOnReadPortRules write SetFOnReadPortRules;   {Event when TPortMetrics needs to read external Port Rules preference}
    property OnWritePortRules : TWritePortRulesEvent read FOnWritePortRules write FOnWritePortRules;  {Event when TPortMetrics needs to write external Port Rules preference}
    property Port[Index: Integer] : String index 0 read GetPort;                                      {List of ports' full descriptions (affected by state of FFiltered and FScannable)}
    property PortID[Index: Integer] : String index 1 read GetPort;                                    {List of Port IDs (affected by state of FFiltered and FScannable)}
    property PortDesc[Index: Integer] : String index 2 read GetPort;                                  {List of Port Descriptions (affected by state of FFiltered and FScannable)}
    property Present[Index: Integer] : Boolean index 0 read GetProperty;                              {True = port is currently present in system, False = port is not present in system}
    property Included[Index: Integer] : Boolean index 1 read GetProperty;                             {True = included, False = not included}
    property InclExclByRule[Index: Integer] : Boolean index 3 read GetProperty;                       {True = included or excluded by a rule.  Read InclExclRuleIndex to get index of that rule.}
    property InclExclRuleIndex[Index: Integer] : Integer read GetRuleIdx;                             {Index of include/exclude rule that was applied (if InclExclByRule = True).}
    property IsSortOrderRule[Index: Integer] : Boolean index 0 read GetRuleType;                      {True = rule at Index is a Sort-Order Port ID rule (may also be an Include/Exclude Port ID rule), False otherwise}
    property IsPortIDRule[Index: Integer] : Boolean index 1 read GetRuleType;                         {True = rule at Index is a Port ID rule, False otherwise}
    property IsPortDescRule[Index: Integer] : Boolean index 2 read GetRuleType;                       {True = rule at Index is a Port Description rule, False otherwise}
    property PortsExcluded : Integer read FPortsExcluded;                                             {Number of "present" ports excluded via rules}
    property Rules : TStrings read FRules;                                                            {List of rules}
    property RuleString: String read GenerateSearchRuleString write ParseSearchRuleString;            {List of rules in single-string format}
    property RuleStringSort: String read GetSortRulesString;                                          {List of rules for port sorting only, in single-string format}
    property RuleStringID: String read GetIDRulesString;                                              {List of rules for Port IDs only, in single-string format}
    property RuleStringDesc: String read GetDescRulesString;                                          {List of rules for Port Descriptions only, in single-string format}
    property Scannable : Boolean read FScannable write SetScannable;                                  {True = FPorts/FPortIDs/FPortDescs contains only ports that are currently scannable, False = contains mixture of scannable and possibly non-scannable ports}
    function EscapeDesc(Description: String): String;                                                 {Returns Description string with properly "escaped" special characters, if they appear}
    function UnEscapeDesc(Description: String): String;                                               {Returns Description string without "escaped" special characters, if they appear}
    function NeedEscapeDesc(Description: String): Integer;                                            {Returns index of character in Description that needs escaping; 0 if none}
    function IndexOfPortID(PortID: String): Integer;                                                  {Returns index of port with PortID}
    function IndexOfPortDesc(MatchDesc: String; AfterIdx: Integer = -1): Integer;                     {Returns index of port matching MatchDesc, starting AfterIdx}
    function RuleMatchCount(MatchDesc: String): Integer;                                              {Returns number of ports matching MatchDesc}
    procedure Refresh;                                                                                {Update entire dataset (refreshes port list from O.S. and filters against Rules)}
  end;                                                                                                 
                                                                                                       

  {TPHintWindow object}                                                                                
  TPCanHintEvent = function (Sender: TObject; ItemID: Integer; var Text: String): Boolean of object;

  TPHintWindow = class(THintWindow)
  private
    FCentered   : Boolean;                                     {True = multi-line (CR/LF delimited) hints are center-justified within the hint window}
    FControl    : TControl;                                    {The control that this hint object belongs to}
    FItemID     : Integer;                                     {User-definable value for current item being hinted}
    FNoItemID   : Integer;                                     {User-definable value indicating no item to hint}
    FShowing    : Boolean;                                     {True if hint currently visible, False otherwise}
    FPos        : TPoint;                                      {Upper-Left coordinate of hint window}
    FCanHint    : TPCanHintEvent;                              {Event that occurs when hint object needs to know if it should provide a hint for FItemID}
    FHidePause  : Integer;                                     {Delay before displayed hint is hidden}
    FShowPause  : Integer;                                     {Delay before hidden hint is displayed}
    FShortPause : Integer;                                     {Delay before next hint is displayed (for quick transitions between two displayed hints)}
    FMaxWidth   : Integer;                                     {Maximum width of hint window}
    FTimer      : TTimer;                                      {The delay timer for hint appearance and disappearance}
    function  CanHint(var NewText: String): Boolean;
    procedure DoShowHint(NewInterval: Integer = 0);
    procedure DoHideHint(NewInterval: Integer = 0);
    procedure TriggerShowHint(Sender: TObject);
    procedure TriggerHideHint(Sender: TObject);
    procedure SetControl(Value: TControl);
    procedure SetNoItemID(Value: Integer);
    procedure SetPos(Value: TPoint);
  public
    property Centered : Boolean read FCentered write FCentered;
    property Control : TControl read FControl write SetControl;
    property HidePause : integer read FHidePause write FHidePause;
    property ItemID : Integer read FItemID write FItemID;
    property MaxWidth : Integer read FMaxWidth write FMaxWidth;
    property NoItemID : Integer read FNoItemID write SetNoItemID;
    property OnCanHint : TPCanHintEvent read FCanHint write FCanHint;
    property Position : TPoint read FPos write SetPos;
    property ShortPause : integer read FShortPause write FShortPause;
    property Showing : Boolean read FShowing;
    property ShowPause : integer read FShowPause write FShowPause;
    property Text;
    constructor Create(Control: TControl); reintroduce;
    destructor  Destroy; override;
    procedure DisableHint(ClearID: Boolean);
    procedure HideHint;
    procedure UpdateHintMetrics(ItemID: Integer; X, Y: Integer);
  end;

var
  PortListForm    : TPortListForm;                              {Main form (Port List)}
  WinPos          : PWindowPlacement;                           {Structure to describe window position status}
  COM             : TPortMetrics;
  SelRow          : Integer;                                    {Current selected Row (valid during mouse-down and popup menu only)}
  OldSelRow       : Integer;                                    {Previous selected row (valid after mouse-up or popup menu selection)}
  UndoRules       : String;                                     {Copy of previous rules before last change}
  RedoRules       : String;                                     {Copy of previous rules before last undo}
  Modified        : Boolean;                                    {Flag indicating a change was made}
  OldCOMFiltered  : Boolean;                                    {State of COM's Filtered flag upon form show}
  OldCOMScannable : Boolean;                                    {State of COM's Scannable flag upon form show}

const
  {WM_DEVICECHANGE message IDs and device types for detecting serial port changes (USB port add/remove)}
  DBT_DEVICEARRIVAL = $8000;
  DBT_DEVICEREMOVECOMPLETE = $8004;
  DBT_DEVTYP_PORT = $3;
  {Other Port List unit constants}
  XEdgeOffset = 5;   {Offset of text from left edge of cell when left justified}
  IncludeAll = '+*';
  ExcludeAll = '-*';
  IncludePort = 'Include Port ';
  ExcludePort = 'Exclude Port ';
  RemovePort = 'Remove Port '{Rule};
  Purpose1 = 'Serial ports matching the description below will be ';
  Purpose2 = ' the search.';
  Included = 'included in';
  Excluded = 'excluded from';

implementation

{$R *.dfm}

{##############################################################################}
{##############################################################################}
{########################## Miscellaneous Routines ############################}
{##############################################################################}
{##############################################################################}

function StrToInt(Str: String): Int64;
{Convert String to 64-bit Integer. If integer value is preceeded by non-digit data, searches until it finds the first valid
digit, then converts until the next invalid digit or end of string.}
var
  Idx   : Integer;
begin
  while (length(Str) > 0) and not (Str[1] in ['0'..'9', '-']) do delete(Str, 1, 1);
  Val(Str, Result, Idx);
end;

{##############################################################################}
{##############################################################################}
{########################## TPortListForm Routines ############################}
{##############################################################################}
{##############################################################################}

{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooo Event Routines oooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}

procedure TPortListForm.WMDeviceChange(var Msg: TDevMsg);
{Handle WM_DEVICECHANGE message.  If it's a Port Arrive or Port Remove message, extract Port ID}
var
  Idx, Len, Step : Integer;
  Chr            : Byte;
  ID             : String;
begin
  if ((Msg.Event = DBT_DEVICEARRIVAL) or (Msg.Event = DBT_DEVICEREMOVECOMPLETE)) and (Msg.Header^.dbch_devicetype = DBT_DEVTYP_PORT) then
    begin {If event is a device arrival or removal, and message structure is a port device type...}
    {Extract Port ID string}
    ID := '';
    Idx := 0;
    Len := Msg.Port^.dbcp_size - 12;                              {Get ID length; structure length minus leading elements}
    Step := 1 + ord(PByteArray(@(Msg.Port^.dbcp_name))[1] = 0);   {Get Step size; 1 if ANSI string, 2 if WideString}
    Chr := byte(Msg.Port^.dbcp_name);                             {Get first char}
    while (Chr > 0) and (Idx < Len) do                            {Loop while char > 0 and haven't reached edge of structure}
      begin
      ID := ID + Char(Chr);                                       {Add char to ID}
      inc(Idx, Step);                                             {Move to next char}
      Chr := PByteArray(@(Msg.Port^.dbcp_name))[Idx];
      end;
    {If ID valid, update display}
    if (length(ID) > 3) and (leftstr(uppercase(ID), 3) = 'COM') and (ID[4] in ['0'..'9']) then
      begin
      COM.DeviceChanged(TPortEvent(ifthen(Msg.Event = DBT_DEVICEARRIVAL, ord(peAdded), ord(peRemoved))), ID);
      if Showing then UpdateDisplay(True);                        {If form showing, update the display (without refreshing serial port list since we've already alerted the COM object)}
      end;
    end;
  inherited;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FormCreate(Sender: TObject);
begin
  {Create Port Description Form}
  FPortDescForm := TPortDescForm.Create(self);
  {Initialize Custom Port List Hints}
  FCPLHint := TPHintWindow.Create(CustomPortList);
  FCPLHint.NoItemID := -1;
  FCPLHint.OnCanHint := CustomPortListCanHint;
  FCPLHint.ShowPause := trunc(Application.HintPause*1.5);
  FCPLHint.ShortPause := 50;
  FCPLHint.Centered := True;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FormDestroy(Sender: TObject);
begin
  FCPLHint.Free;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FormShow(Sender: TObject);
begin
  {Force CustomPortList canvas' font to be same as CustomPortList's font (otherwise, first run of SizeComponents malfunctions because Canvas initializes with default font "MS Sans Serif")}
  CustomPortList.Canvas.Font := CustomPortList.Font;
  {Set Header}
  HeaderLabel.Caption := 'The '+FDeviceName+' loading process will scan these serial ports, in the order shown.';
  {Save Filtered and Scannable flag states and set properly for this form}
  OldCOMFiltered := COM.Filtered;
  OldCOMScannable := COM.Scannable;
  COM.Filtered := True;
  COM.Scannable := False;
  {Draw entire display}
  UpdateDisplay;
  {Reset Modified flag}
  Modified := False;
  {Make sure form is at reasonably displayable coordinates}
  EnsureWindowDisplayable(self);
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
  Response : Cardinal;
begin
  {Reset Filtered and Scannable flag states}
  COM.Filtered := OldCOMFiltered;
  COM.Scannable := OldCOMScannable;
  {Process close request}
  if not Modified then exit;
  {Prompt user}
  messagebeep(MB_ICONWARNING);
  Response := MessageDlg('Discard all changes to Serial Port Search List?', mtConfirmation, [mbYES, mbNO], 0);
  if Response = mrCancel then
    begin  {Abort closure}
    Action := caNone;
    exit;
    end;
  if Response = mrYES then {Okay to discard changes and close?}
    begin  {Discard changes}
    COM.DoReadPortRules;  {Have TPortMetrics re-read external port rules preference}
    exit;
    end
  else
  FinalizeModifications;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.AcceptButtonClick(Sender: TObject);
begin
  FinalizeModifications;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.RestoreDefaultsButtonClick(Sender: TObject);
var
  Response : Cardinal;
begin
  {Prompt user}
  messagebeep(MB_ICONWARNING);
  Response := MessageDlg('Discard all custom rules and reset them to defaults?', mtConfirmation, [mbYES, mbNO], 0);
  if Response = mrYes then {Okay to discard custom rules and reset to defaults}
    begin
    Undo(False);
    COM.DoReadPortRules(True);
    UpdateDisplay;
    end;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
{Custom draw cell of Custom Port List}
var
  XOff, YOff    : Integer;
  W, H          : Integer;
  OldFontColor  : TColor;
  OldBrushColor : TColor;
begin
  with CustomPortList do
    begin
    {Save original colors}
    OldFontColor := Canvas.Font.Color;
    OldBrushColor := Canvas.Brush.Color;
    {Determine and set new colors}
    if ARow = 0 then
      begin {Header row}
      Canvas.Font.Color := clWindowText;
      Canvas.Brush.Color := clBtnFace;
      Canvas.Font.Style := [fsBold];
      end
    else
      begin {Port row}
      if (ARow <> SelRow) then
        begin {Not selected row}
        if COM.Excluded[ARow-1] {Port Excluded} then Canvas.Font.Color := clGrayText else Canvas.Font.Color := clWindowText;
        Canvas.Brush.Color := clWindow;
        end
      else    {Selected row}
        begin
        Canvas.Font.Color := clHighlightText;
        Canvas.Brush.Color := clHighlight;
        end;
      if COM.Excluded[ARow-1] {Port Excluded} then Canvas.Font.Style := [fsItalic] else Canvas.Font.Style := [];
      end;

    {Draw cell background}
    Canvas.FillRect(Rect);

    {Determine text width, height, and offset within cell}
    W := Canvas.TextWidth(CustomPortList.Cells[ACol, ARow]);
    H := Canvas.TextHeight(CustomPortList.Cells[ACol, ARow]);
    YOff := (Rect.Bottom - Rect.Top) div 2 - (H div 2);
    XOff := ifthen((ARow = 0) or (ACol <> 1), (Rect.Right - Rect.Left) div 2 - (W div 2), 5) - ifthen(fsItalic in Canvas.Font.Style, 2, 0);

    {Draw text}
    Canvas.Brush.Style := bsClear;
    Canvas.TextOut(Rect.Left+XOff, Rect.Top+YOff, CustomPortList.Cells[ACol, ARow]);

    {Restore original colors}
    Canvas.Font.Color := OldFontColor;
    Canvas.Brush.Color := OldBrushColor;
    end;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{Determine selected row and repaint Custom Port List}
var
  Col, Row : Integer;
begin
  FCPLHint.DisableHint(True);
  CustomPortList.MouseToCell(X, Y, Col, Row);
  SelRow := Row;
  CustomPortList.Repaint;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
{Mouse moved over port list, determine if over port item and process hint}
var
  Col, Row : Integer;
begin
  CustomPortList.MouseToCell(X, Y, Col, Row);         {Get index of row under mouse}
  FCPLHint.UpdateHintMetrics(max(-1, Row-1), X, Y);    {Update hint metrics, adjusting row index to indicate port item}
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{Deselect row and repaint Custom Port List}
begin
  OldSelRow := SelRow;
  SelRow := 0;
  CustomPortList.Repaint;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
{Process row sorting event.}
var
  Idx : Integer;
  Str : String;
begin
  Str := '';
  for Idx := 1 to CustomPortList.RowCount-1 do Str := Str + '(' + inttostr(strtoint(CustomPortList.Cells[0, Idx])) + '),';
  Undo(False);
  COM.RuleString := Str + COM.RuleString;
  UpdateDisplay;
end;

{------------------------------------------------------------------------------}

function TPortListForm.CustomPortListCanHint(Sender: TObject; ItemID: Integer; var Text: String): Boolean;
{Process custom port list can-hint event}
begin
  Result := True;
  Text := ifthen(COM.Included[ItemID], 'Included', 'Excluded') + ' by ';
  if COM.InclExclByRule[ItemID] then
    Text := Text + ifthen(COM.IsPortIDRule[COM.InclExclRuleIndex[ItemID]], 'Port ID', 'Port Description') + ' rule'
  else
    Text := Text + 'default';
  if not COM.Present[ItemID] then Text := Text + #$D#$A+ '(This port is currently absent)';
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.CustomPortListContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
{Configure Custom Port List shortcut menu before it appears.}
begin
  FCPLHint.DisableHint(True);
  if SelRow > 0 then
    begin {Port row selected}
    IncludePortItem.Visible := COM.Present[SelRow-1] and COM.Excluded[SelRow-1];
    IncludePortItem.Caption := ifthen(COM.InclExclByRule[SelRow-1] and COM.IsPortIDRule[COM.InclExclRuleIndex[SelRow-1]], 'Re-'+IncludePort, IncludePort) + '(' + COM.PortID[SelRow-1] + ')';
    ExcludePortItem.Visible := COM.Present[SelRow-1] and  COM.Included[SelRow-1];
    ExcludePortItem.Caption := ifthen(COM.InclExclByRule[SelRow-1] and COM.IsPortIDRule[COM.InclExclRuleIndex[SelRow-1]], 'Re-'+ExcludePort, ExcludePort) + '(' + COM.PortID[SelRow-1] + ')';
    RemovePortRuleItem.Visible := not COM.Present[SelRow-1];
    RemovePortRuleItem.Caption := RemovePort + '(' + COM.PortID[SelRow-1] + ')' + ' Rule';
    FilterPortsByDescriptionItem.Visible := COM.Present[SelRow-1] and not COM.InclExclByRule[SelRow-1];
    EditPortDescriptionRuleItem.Visible := COM.Present[SelRow-1] and (COM.InclExclByRule[SelRow-1] and COM.IsPortDescRule[COM.InclExclRuleIndex[SelRow-1]]);
    RemovePortDescriptionRuleItem.Visible := COM.Present[SelRow-1] and (COM.InclExclByRule[SelRow-1] and COM.IsPortDescRule[COM.InclExclRuleIndex[SelRow-1]]);
    end
  else
    begin {Header row (0) or blank area selected}
    IncludePortItem.Visible := False;
    ExcludePortItem.Visible := False;
    RemovePortRuleItem.Visible := False;
    FilterPortsByDescriptionItem.Visible := False;
    EditPortDescriptionRuleItem.Visible := False;
    RemovePortDescriptionRuleItem.Visible := False;
    end;
  IncludeAllPortsByDefaultItem.Visible := COM.GlobalExclude;
  ExcludeAllPortsByDefaultItem.Visible := COM.GlobalInclude;
  UndoLastChangeItem.Visible := UndoRules <> '';
  RedoPriorChangeItem.Visible := RedoRules <> '';
  ClearAllRulesItem.Visible := COM.Rules.Count > 1;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.PopupMenuItemClick(Sender: TObject);
{Process Port List popup menu item selection.}
var
  Temp    : String;
  Result  : Integer;
  WinRect : TRect;
begin
  case TMenuItem(Sender).Tag of
    10, 11 : if OldSelRow > 0 then                                                                                      {Include/Exclude port by ID}
               begin
               Undo(False);
               if COM.InclExclByRule[OldSelRow-1] and COM.IsPortIDRule[COM.InclExclRuleIndex[OldSelRow-1]] then         {Need to re-include/exclude? (ie: need to remove Include/Exclude Port ID rule}
                 begin
                 Temp := COM.Rules[COM.InclExclRuleIndex[OldSelRow-1]];
                 if Temp[1] <> '(' then                                                                                 {  If not sort-order rule}
                   COM.Rules.Delete(COM.InclExclRuleIndex[OldSelRow-1])                                                 {    delete rule}
                 else
                   begin                                                                                                {  else}
                   delete(Temp, 2, 1);                                                                                  {    remove include/exclude char}
                   COM.Rules[COM.InclExclRuleIndex[OldSelRow-1]] := Temp;                                               {    and retain sort-order rule}
                   end;
                 end
               else                                                                                                     {Else, need to include/exclude}
                 COM.RuleString := COM.RuleString + ',' + ifthen(TMenuItem(Sender).Tag and 1 = 1, '-', '+') + inttostr(strtoint(COM.PortID[OldSelRow-1]));
               end;
    12, 22 : begin                                                                                                      {Remove Port ID or Port Description rule}
             if TMenuItem(Sender).Tag = 22 then
               begin
               Result := COM.RuleMatchCount(COM.Rules[COM.InclExclRuleIndex[OldSelRow-1]]);
               Temp := 'Delete the Port Description rule affecting this port?'+#$D#$A#$D#$A+'(This rule is ';
               Temp := Temp + ifthen(Result = 1, 'not currently affecting any other ports.)', 'currently affecting '+inttostr(Result-1)+' other port'+ifthen(Result > 2, 's', '')+' as well.)');
               messagebeep(MB_ICONWARNING);
               Result := messagedlg(Temp, mtConfirmation, [mbYes, mbNo], 0);
               end
             else
               Result := mrYes;
             if Result = mrYes then
               begin {Okay to delete}
               Undo(False);
               COM.Rules.Delete(COM.InclExclRuleIndex[OldSelRow-1]);
               end;
             end;
    20, 21 : begin                                                                                                      {Filter/Edit Port by Description}
             {Position Port Description window below selected row}
             WinRect := CustomPortList.CellRect(1, OldSelRow);
             WinRect.Top := WinRect.Bottom + 3;
             WinRect.Left := WinRect.Left + 3;
             WinRect.TopLeft := CustomPortList.ClientToScreen(WinRect.TopLeft);
             WinRect.BottomRight := point(WinRect.Left+FPortDescForm.Width, WinRect.Top+FPortDescForm.Height);
             {Ensure window will be visible}
             WinRect := GetDisplayableWindowPosition(WinRect);
             {Set to position}
             WinPos.flags := 0;
             WinPos.ptMaxPosition := Point(0,0);
             WinPos.rcNormalPosition := WinRect;
             WinPos.showCmd := SW_HIDE;
             SetWindowPlacement(FPortDescForm.Handle, WinPos);
             {Prompt user with Port Description window}
             if FPortDescForm.ShowModal(COM.InclExclRuleIndex[OldSelRow-1], COM.PortDesc[OldSelRow-1]) = mrOK then
               begin
               Undo(False);
               Temp := ifthen(FPortDescForm.FilterComboBox.ItemIndex = 1, '-', '+') + FPortDescForm.DescriptionEdit.Text;
               if COM.InclExclRuleIndex[OldSelRow-1] = -1 then COM.RuleString := COM.RuleString + ',' + Temp else COM.Rules[COM.InclExclRuleIndex[OldSelRow-1]] := Temp;
               end;
             end;
    30, 31 : begin                                                                                                      {Include/Exclude all ports by default}
             Undo(False);
             COM.RuleString := ifthen(TMenuItem(Sender).Tag and 1 = 1, ExcludeAll, IncludeAll) + ',' + COM.RuleStringSort;
             end;
    40, 41 : if TMenuItem(Sender).Tag and 1 = 1 then Redo else Undo(True);                                              {Redo/Undo operation}
    50     : begin                                                                                                      {Clear all rules}
             Undo(False);
             COM.RuleString := '';
             end;
  end;
  UpdateDisplay;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{Process special keys}
begin
  if (Shift = [ssCtrl]) and (uppercase(char(Key)) = 'Z') then
    begin
    Key := 0;
    if UndoRules <> '' then Undo(True) else Redo;
    end;
end;

{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooo Non-Event Routines oooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}

function TPortListForm.GetDisplayableWindowPosition(Bounds: TRect): TRect;
{Return window bounds that ensures at least 50% of the window's title bar is visible on a display.
 Bounds must be the window's current left, top, right and bottom coordinates.
 If those coordinates result in the window's title bar being more than 50% outside any visible desktop, a new bounding rectangle is returned that
 centers the window in the nearest monitor, otherwise, the same bounding rectangle coordinates are returned.}
var
  Monitor     : HMonitor;
  MonitorInfo : TMonitorInfo;
  TitleCenter : TPoint;
  Girth       : TPoint;         {Bounding rectangle's width and height}
begin
  Result := Bounds;
  {Find window's girth (width and height) and the center of it's title bar}
  Girth := point(Bounds.Right - Bounds.Left, Bounds.Bottom - Bounds.Top);
  TitleCenter := point(Bounds.Left + (Girth.X div 2), Bounds.Top + getsystemmetrics(SM_CYFRAME) + (getsystemmetrics(SM_CYCAPTION) div 2));
  {Find monitor that is displaying title bar's center point}
  Monitor := MonitorFromPoint(TitleCenter, MONITOR_DEFAULTTONULL);
  if Monitor <> 0 then exit;
  {No monitor is displaying that point?  Find the nearest one and center the Bounds area within it}
  Monitor := MonitorFromPoint(TitleCenter, MONITOR_DEFAULTTONEAREST);
  MonitorInfo.cbSize := sizeof(TMonitorInfo);
  if not GetMonitorInfo(Monitor, @MonitorInfo) then exit;
  with MonitorInfo.rcMonitor do
    begin
    Result.TopLeft := point(max(Left, Left + ((Right - Left) div 2) - (Girth.X div 2)), max(Top, Top + ((Bottom - Top) div 2) - (Girth.Y div 2)));
    Result.BottomRight := point(Result.Left + Girth.X, Result.Top + Girth.Y);
    end;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.EnsureWindowDisplayable(Window: TForm);
{Ensure window is reasonably within displayable coordinates.  Moves window if necessary.}
begin
  GetWindowPlacement(Window.Handle, WinPos);
  WinPos.rcNormalPosition := GetDisplayableWindowPosition(Window.BoundsRect);
  if not (PointsEqual(WinPos.rcNormalPosition.TopLeft, Window.BoundsRect.TopLeft) and PointsEqual(WinPos.rcNormalPosition.BottomRight, Window.BoundsRect.BottomRight)) then
    SetWindowPlacement(Window.Handle, WinPos);
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.SizeComponents;
{Size Custom Port List in relation to the data to display}
var
  Idx          : Integer;
  TotWidth     : Integer;
  OldFontStyle : TFontStyles;
  GridH, GridW,
  EdgeSize     : Integer;
const
  MinRows = 1 + 2;   {Minimum of 1 header row and 2 port rows}
  MaxRows = 1 + 10;  {Maximum of 1 header row and 10 port rows}
begin
  {Save current font style and clear it}
  OldFontStyle := CustomPortList.Canvas.Font.Style;
  CustomPortList.Canvas.Font.Style := [];
  {Determine row heights and widths}
  CustomPortList.DefaultRowHeight := CustomPortList.Canvas.TextHeight('X')+4;
  GridH := CustomPortList.GridLineWidth*ord(goFixedHorzLine in CustomPortList.Options);
  GridW := CustomPortList.GridLineWidth*ord(goFixedVertLine in CustomPortList.Options);
  EdgeSize := CustomPortList.GridLineWidth*2*ord(CustomPortList.BorderStyle = bsSingle)*2*ord(CustomPortList.Ctl3D)+1*ord(not CustomPortList.Ctl3D and (CustomPortList.BorderStyle = bsSingle));
  CustomPortList.Height := min(MaxRows, max(MinRows, COM.Count+1))*(CustomPortList.DefaultRowHeight + GridH) + EdgeSize;
  CustomPortList.ColWidths[0] := CustomPortList.Canvas.TextWidth('COM9999')+XEdgeOffset;
  CustomPortList.ColWidths[2] := CustomPortList.Canvas.TextWidth('PRESENT')+XEdgeOffset;
  CustomPortList.ColWidths[3] := 0;
  CustomPortList.ColWidths[1] := max(CustomPortList.ColWidths[1], InstructionPanel.ClientToScreen(ClickAndDragLabel.BoundsRect.TopLeft).X -
                                                                  self.ClientToScreen(CustomPortList.BoundsRect.TopLeft).X +
                                                                  ClickAndDragLabel.Canvas.TextWidth(ClickAndDragLabel.Caption) -
                                                                  CustomPortList.ColWidths[0] - CustomPortList.ColWidths[2]);
  for Idx := 0 to COM.Count-1 do CustomPortList.ColWidths[1] := max(CustomPortList.ColWidths[1], CustomPortList.Canvas.TextWidth(COM.PortDesc[Idx])+XEdgeOffset*2);
  {Vertically position lower components}
  InstructionPanel.Top := CustomPortList.Top + CustomPortList.Height + 15;
  RestoreDefaultsButton.Top := InstructionPanel.Top + InstructionPanel.Height + 15;
  AcceptButton.Top := RestoreDefaultsButton.Top;
  CancelButton.Top := RestoreDefaultsButton.Top;
  ClientHeight := CancelButton.Top + CancelButton.Height + CustomPortList.Left;
  {Size widths of controls}
  TotWidth := 0;
  for Idx := 0 to CustomPortList.ColCount-2 do inc(TotWidth, CustomPortList.ColWidths[Idx] + GridW);
  CustomPortList.Width := TotWidth + EdgeSize + ifthen(CustomPortList.VisibleRowCount+1 < CustomPortList.RowCount, GetSystemMetrics(SM_CXVSCROLL), 0);
  HeaderLabel.Width := CustomPortList.Width;
  ItalicHeaderLabel.Width := CustomPortList.Width;
  InstructionPanel.Width := CustomPortList.Width;
  ClientWidth := CustomPortList.Left * 2 + CustomPortList.Width;
  {Horizontally position button}
  RestoreDefaultsButton.Left := CustomPortList.Left;
  AcceptButton.Left := ClientWidth - CancelButton.Width - CustomPortList.Left*2 - AcceptButton.Width;
  CancelButton.Left := ClientWidth - CancelButton.Width - CustomPortList.Left;
  {Restore font style}
  CustomPortList.Canvas.Font.Style := OldFontStyle;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.UpdateDisplay(NoRefresh: Boolean = False);
{If NoRefresh, don't refresh COM port list.}
var
  Idx : Integer;
begin
  if not NoRefresh then COM.Refresh;
  CustomPortList.RowCount := COM.Count+1;
  OldSelRow := SelRow;
  SelRow := 0;
  SizeComponents;
  CustomPortList.Cells[0, 0] := 'Port ID';
  CustomPortList.Cells[1, 0] := 'Port Description';
  CustomPortList.Cells[2, 0] := 'Present';
  for Idx := 0 to COM.Count-1 do
    begin {For every port found, add it to the Custom Port List}
    CustomPortList.Cells[0, Idx+1] := COM.PortID[Idx];
    CustomPortList.Cells[1, Idx+1] := COM.PortDesc[Idx];
    CustomPortList.Cells[2, Idx+1] := ifthen(COM.Present[Idx], 'Yes', 'No');
    end;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.Undo(PerformUndo: Boolean);
{Record undo operation, or PerformUndo operation.}
begin
  if not PerformUndo then
    begin                              {Record undo operation}
    Modified := True;
    UndoRules := COM.RuleString;
    RedoRules := '';
    end
  else
    begin                              {Perform undo operation}
    if UndoRules = '' then exit;
    Modified := True;
    RedoRules := COM.RuleString;
    COM.RuleString := UndoRules;
    UndoRules := '';
    UpdateDisplay;
    end;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.Redo;
{Perform Redo operation}
begin
  if RedoRules = '' then exit;
  Modified := True;
  UndoRules := COM.RuleString;
  COM.RuleString := RedoRules;
  RedoRules := '';
  UpdateDisplay;
end;

{------------------------------------------------------------------------------}

procedure TPortListForm.FinalizeModifications;
begin
  COM.DoWritePortRules;
  Modified := False;
end;

{##############################################################################}
{##############################################################################}
{########################## TPortDescForm Routines ############################}
{##############################################################################}
{##############################################################################}

{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooo Event Routines oooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}

constructor TPortDescForm.Create(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  {Set form properties}
  BorderIcons := [biSystemMenu];
  BorderStyle := bsDialog;
  Caption := 'Filter Port(s) By Description';
  ClientHeight := 305;
  ClientWidth := 397;
  Color := clBtnFace;
  Font.Charset := DEFAULT_CHARSET;
  Font.Color := clWindowText;
  Font.Height := -11;
  Font.Name := 'Arial';
  Font.Style := [];
  OldCreateOrder := False;
  Scaled := False;
  OnActivate := FormActivate;
  PixelsPerInch := 96;
  ParentFont := False;
  {Purpose Label}
  PurposeLabel := TLabel.Create(self);
  with PurposeLabel do
    begin
    Parent := self;
    Left := 16;
    Top := 16;
    Width := 328;
    Height := 32;
    Alignment := taCenter;
    Caption := 'Serial ports matching the description below will be included in the search.';
    Font.Charset := DEFAULT_CHARSET;
    Font.Color := clWindowText;
    Font.Height := -13;
    Font.Name := 'Arial';
    Font.Style := [fsBold];
    ParentFont := False;
    WordWrap := True;
    end;
  {Filter Combobox}
  FilterComboBox := TComboBox.Create(self);
  with FilterComboBox do
    begin
    Parent := self;
    Left := 16;
    Top := 64;
    Width := 65;
    Height := 22;
    DropDownCount := 2;
    ParentFont := True;
    ItemHeight := 14;
    ItemIndex := 0;
    TabOrder := 0;
    Text := 'Include';
    OnChange := FilterComboBoxChange;
    Items.Text := 'Include'+#$D#$A+'Exclude';
    end;
  {Description Label}
  DescriptionLabel := TLabel.Create(self);
  with DescriptionLabel do
    begin
    Parent := self;
    Left := 86;
    Top := 67;
    Width := 130;
    Height := 14;
    Caption := 'ports matching description:';
    ParentFont := True;
    end;
  {Description Edit}
  DescriptionEdit := TEdit.Create(self);
  with DescriptionEdit do
    begin
    Parent := self;
    Left := 220;
    Top := 64;
    Width := 157;
    Height := 22;
    TabOrder := 1;
    OnKeyUp := DescriptionEditKeyUp;
    end;
  {Instruction Panel}
  InstructionPanel := TPanel.Create(self);
  with InstructionPanel do
    begin
    Parent := self;
    Left := 16;
    Top := 97;
    Width := 361;
    Height := 144;
    BevelOuter := bvNone;
    Ctl3D := True;
    ParentCtl3D := False;
    TabOrder := 4;
    {Tips Label}
    TipsLabel := TLabel.Create(InstructionPanel);
    with TipsLabel do
      begin
      Parent := InstructionPanel;
      Left := 0;
      Top := 9;
      Width := 30;
      Height := 16;
      Caption := 'Tips:';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -13;
      Font.Name := 'Arial';
      Font.Style := [fsBold];
      ParentFont := False;
      end;
    {Bullets}
    Bullet1 := TShape.Create(InstructionPanel);
    with Bullet1 do
      begin
      Parent := InstructionPanel;
      Left := 20;
      Top := 42;
      Width := 5;
      Height := 5;
      Brush.Color := clBlack;
      Shape := stCircle;
      end;
    Bullet2 := TShape.Create(InstructionPanel);
    with Bullet2 do
      begin
      Parent := InstructionPanel;
      Left := 20;
      Top := 67;
      Width := 5;
      Height := 5;
      Brush.Color := clBlack;
      Shape := stCircle;
      end;
    Bullet3 := TShape.Create(InstructionPanel);
    with Bullet3 do
      begin
      Parent := InstructionPanel;
      Left := 20;
      Top := 91;
      Width := 5;
      Height := 5;
      Brush.Color := clBlack;
      Shape := stCircle;
      end;
    Bullet4 := TShape.Create(InstructionPanel);
    with Bullet4 do
      begin
      Parent := InstructionPanel;
      Left := 20;
      Top := 115;
      Width := 5;
      Height := 5;
      Brush.Color := clBlack;
      Shape := stCircle;
      end;
    {General/Specific Label}
    GenSpecLabel := TLabel.Create(InstructionPanel);
    with GenSpecLabel do
      begin
      Parent := InstructionPanel;
      Left := 30;
      Top := 37;
      Width := 309;
      Height := 15;
      Caption := 'Description may be as general or as specific as desired.';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -12;
      Font.Name := 'Arial';
      Font.Style := [];
      ParentFont := False;
      end;
    {Asterisk Label}
    AsteriskLabel := TLabel.Create(InstructionPanel);
    with AsteriskLabel do
      begin
      Parent := InstructionPanel;
      Left := 30;
      Top := 62;
      Width := 309;
      Height := 15;
      Caption := 'Asterisks (*) can be used as wildcards.';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -12;
      Font.Name := 'Arial';
      Font.Style := [];
      ParentFont := False;
      end;
    {DescCase Label}
    DescCaseLabel := TLabel.Create(InstructionPanel);
    with DescCaseLabel do
      begin
      Parent := InstructionPanel;
      Left := 30;
      Top := 86;
      Width := 309;
      Height := 15;
      Caption := 'Description is not case sensitive.';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -12;
      Font.Name := 'Arial';
      Font.Style := [];
      ParentFont := False;
      end;
    {Override Label}
    OverrideLabel := TLabel.Create(InstructionPanel);
    with OverrideLabel do
      begin
      Parent := InstructionPanel;
      Left := 30;
      Top := 110;
      Caption := 'Port Description rules are overriden by Port ID rules whenever both match a port.';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -12;
      Font.Name := 'Arial';
      Font.Style := [];
      ParentFont := False;
      WordWrap := True;
      Width := 309;
      Height := 30;
      end;
    end;
  {Matches Label}
  MatchesLabel := TLabel.Create(self);
  with MatchesLabel do
    begin
    Parent := self;
    Left := 16;
    Top := 268;
    Width := 100;
    Height := 16;
    Caption := 'Matches Found:';
    Font.Charset := DEFAULT_CHARSET;
    Font.Color := clWindowText;
    Font.Height := -13;
    Font.Name := 'Arial';
    Font.Style := [fsBold];
    ParentFont := False;
    end;
  {Match Count Label}
  MatchCountLabel := TLabel.Create(self);
  with MatchCountLabel do
    begin
    Parent := self;
    Left := 124;
    Top := 269;
    Width := 7;
    Height := 15;
    Alignment := taCenter;
    Caption := '0';
    Font.Charset := DEFAULT_CHARSET;
    Font.Color := clWindowText;
    Font.Height := -12;
    Font.Name := 'Arial';
    Font.Style := [fsBold];
    ParentFont := False;
    WordWrap := True;
    end;
  {Okay Button}
  OkayButton := TButton.Create(self);
  with OkayButton do
    begin
    Parent := self;
    Left := 216;
    Top := 264;
    Width := 75;
    Height := 25;
    Caption := '&OK';
    Default := True;
    ModalResult := mrOK;
    TabOrder := 2;
    OnClick := OkayButtonClick;
    end;
  {Cancel Button}
  CancelButton := TButton.Create(self);
  with CancelButton do
    begin
    Parent := self;
    Left := 304;
    Top := 264;
    Width := 75;
    Height := 25;
    Caption := '&Cancel';
    Cancel := True;
    ModalResult := mrCancel;
    TabOrder := 3;
    end;
end;

{------------------------------------------------------------------------------}

function TPortDescForm.ShowModal(RuleIdx: Integer; Description: String): Integer;
begin
  if RuleIdx = -1 then
    begin  {No current rule, set to opposite global rule and use provided description}
    FilterComboBox.ItemIndex := ord(COM.GlobalInclude);
    DescriptionEdit.Text := COM.EscapeDesc(Description);
    end
  else
    begin
    FilterComboBox.ItemIndex := ord(COM.Rules[RuleIdx][1] = '-');
    DescriptionEdit.Text := rightstr(COM.Rules[RuleIdx], length(COM.Rules[RuleIdx])-1);
    end;
  UpdatePurpose;
  UpdateMatchCount;
  Result := inherited ShowModal;
end;

{------------------------------------------------------------------------------}

procedure TPortDescForm.FormActivate(Sender: TObject);
begin
  DescriptionEdit.SetFocus;
end;

{------------------------------------------------------------------------------}

procedure TPortDescForm.OkayButtonClick(Sender: TObject);
var
  Idx   : Integer;
  Comma : Boolean;
begin
  Idx := COM.NeedEscapeDesc(DescriptionEdit.Text);
  if Idx > 0 then
    begin
    ModalResult := mrNone;
    messagebeep(MB_ICONERROR);
    Comma := DescriptionEdit.Text[Idx] = ',';
    messagedlg('Invalid Description at character# ' + inttostr(Idx) + '.' + #$D#$A#$D#$A +
            ifthen(Comma, 'Comma characters  ( , )  must be "escaped"  ( \, )', 'Backslash characters  ( \ )  must be escaped  ( \\ )') + '  in the Description.', mtError, [mbOk], 0);
    DescriptionEdit.SetFocus;
    end
  else
    if MatchCount = 0 then
      begin
      messagebeep(MB_ICONWARNING);
      if messagedlg('Port Description rule "' + DescriptionEdit.Text + '" does not match any present serial ports!' + #$D#$A#$D#$A +
                    'If you save this Port Description rule, it may match a serial port in the future, however you will not be able to edit this rule again until such a port is present.' + #$D#$A#$D#$A +
                    'Continue saving this Port Description rule?', mtWarning, [mbYes, mbNo], 0) = mrNo then ModalResult := mrNone;
      end;
end;

{------------------------------------------------------------------------------}

procedure TPortDescForm.FilterComboBoxChange(Sender: TObject);
begin
  UpdatePurpose;
end;

{------------------------------------------------------------------------------}

procedure TPortDescForm.DescriptionEditKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  UpdateMatchCount;
end;

{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooo Non-Event Routines oooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}
{oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo}

procedure TPortDescForm.UpdatePurpose;
begin
  PurposeLabel.Caption := Purpose1 + ifthen(FilterComboBox.ItemIndex = 0, Included, Excluded) + Purpose2;
  PurposeLabel.Width := DescriptionEdit.Left+DescriptionEdit.Width - PurposeLabel.Left;
end;

{------------------------------------------------------------------------------}

procedure TPortDescForm.UpdateMatchCount;
begin
  MatchCount := COM.RuleMatchCount(ifthen(FilterComboBox.ItemIndex = 1, '-', '+') + DescriptionEdit.Text);
  MatchCountLabel.Caption := inttostr(MatchCount);
end;

{##############################################################################}
{##############################################################################}
{########################### TPortMetrics Routines ############################}
{##############################################################################}
{##############################################################################}

{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบ     Private Routines     บบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}

procedure TPortMetrics.DeviceChanged(PortEvent: TPortEvent; PortID: String);
{A device, PortID, was changed as indicated by PortEvent.  This method should only be called by the WM_DEVICECHANGE message handler from the TPortListForm.}
begin
  Refresh;
  PortID := ExtractPortID(PortID); {Clean PortID (if necessary)}
  if assigned(FOnDeviceChange) then FOnDeviceChange(self, PortEvent, PortID);
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.DoReadPortRules(Default: Boolean = False);
{Perform ReadPortRules event.  This reads the external port rules preference.
 If Default = True, this reads the default external port rules preference.}
begin
  if assigned(FOnReadPortRules) then ParseSearchRuleString(FOnReadPortRules(Self, Default));
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.DoWritePortRules;
{Perform WritePortRules event.  This writes the external port rules preference.}
begin
  if assigned(FOnWritePortRules) then FOnWritePortRules(Self, GenerateSearchRuleString);
end;

{------------------------------------------------------------------------------}

{NOTE: The following method is left intact from it's standard form as defined by Parallax to ease its future maintenance, even though doing so makes it not
as efficient as it could be in the context of the TPortMetrics object.}

procedure TPortMetrics.EnumerateComPorts(LongNames: Boolean; var Ports: TStrings);
{Use the Setup Device Installation API to retrieve the installed COM ports (on Win2K, and above, only)}
var
  PortsGUID       : array of TGUID;     {The array of GUIDs for the Ports class}
  DeviceInfo      : TSPDevInfoData;     {The Device Info structure for a particular hardware device}
  DeviceInfoList  : HDEVINFO;           {The list of devices found under the Ports class}
  ReqSize         : Cardinal;           {Required buffer size (value returned by SetupDi API calls}
  Idx, GUIDIdx    : Integer;
  Buffer          : string;             {Temporary string storage}
  COMRef          : string;             {Holds COM port reference ID, if any}

    {----------------------}

    procedure AddIfPortValid;
    {Adds port (in Buffer) to Ports list if it contains a valid COM port identifier reference.
     Port is added in either Long Name or Short Name format.}
    var
      PIdx  : Integer;
      Valid : Boolean;
    begin
      Valid := False;
      COMRef := Buffer;
      try
        repeat
          PIdx := pos('COM', COMRef);     {Possible COM reference found?}
          if PIdx = 0 then exit;          {Exit if not}
          delete(COMRef, 1, PIdx-1);      {Strip leading chars}
          PIdx := 4;                      {Validate reference, or delete it if invalid}
          while (PIdx <= length(COMRef)) and (COMRef[PIdx] in ['0'..'9']) do inc(PIdx);
          if PIdx = 4 then delete(COMRef, 1, 3) else delete(COMRef, PIdx, length(COMRef));
          Valid := PIdx > 4;              {Is valid?}
        until PIdx > length(COMRef);      {Loop if not}
      finally
        if Valid then
          begin
          if LongNames then Ports.Add(Buffer) else Ports.Add(COMRef);
          end;
      end;
    end;

    {----------------------}

begin
  Ports.Clear;
  {Get GUID for 'PORTS'.}
  {First, intentionally pass a nil pointer and size 0 to retrieve the required PortsGUID array size.}
  if SetupDiClassGuidsFromName('PORTS', nil, 0, ReqSize) then raise Exception.Create('Error receiving PortsGUID.  Error is unknown.');
  if GetLastError <> ERROR_INSUFFICIENT_BUFFER then raise Exception.Create('Error receiving PortsGUID.  Error# '+inttostr(GetLastError)+'.');
  {Second, size the array properly and try again}
  setlength(PortsGUID, ReqSize);
  try
    if not SetupDiClassGuidsFromName('PORTS', @PortsGUID[0], length(PortsGUID), ReqSize) then raise Exception.Create('Error receiving PortsGUID.  Error# '+inttostr(GetLastError)+'.');
    {Get list of devices that are "present" and expose an "interface" for class}
    for GUIDIdx := 0 to length(PortsGUID)-1 do
      begin {For each GUID found...}
      DeviceInfoList := SetupDiGetClassDevs(@PortsGUID[GUIDIdx], nil, 0, DIGCF_PRESENT);
      if cardinal(DeviceInfoList) = INVALID_HANDLE_VALUE then raise Exception.Create('Could not create device info list for PortGUID index '+inttostr(GUIDIdx)+'!');
      try {Enumerate all ports}
        Idx := 0;
        DeviceInfo.cbSize := sizeof(TSPDevInfoData);
        while SetupDiEnumDeviceInfo(DeviceInfoList, Idx, DeviceInfo) do
          begin {For each device found, get device's "friendly" name}
          if not SetupDiGetDeviceRegistryProperty(DeviceInfoList, DeviceInfo, SPDRP_FRIENDLYNAME, nil, nil, 0, @ReqSize) and (GetLastError = ERROR_INSUFFICIENT_BUFFER) then
            begin {Received required size value, now set buffer to that size and call again}
            SetLength(Buffer, ReqSize);
            SetupDiGetDeviceRegistryProperty(DeviceInfoList, DeviceInfo, SPDRP_FRIENDLYNAME, nil, @Buffer[1], ReqSize, nil);
            AddIfPortValid;
            end;
          inc(Idx);
          end;
      finally {Free memory used for Device Info List}
        SetupDiDestroyDeviceInfoList(DeviceInfoList);
      end; {try..finally}
      end; {for...}
    finally {Free memory used for PortsGUID}
      setlength(PortsGUID, 0);
    end; {try..finally}
end;

{------------------------------------------------------------------------------}

function TPortMetrics.ExtractPortID(Port: String): String;
{Returns Port ID (COMxxx) from Port string.}
var
  PIdx   : Integer;
begin
  Result := '';
  repeat
    PIdx := pos('COM', uppercase(Port));  {Possible COM reference found?}
    if PIdx = 0 then exit;                {Exit if not}
    delete(Port, 1, PIdx+2);              {Strip leading chars, including 'COM'}
    while (length(Port) > 0) and (Port[1] in ['0'..'9']) do
      begin {Parse numerics}
      Result := Result + Port[1];
      delete(Port, 1, 1);
      end;
  until Result <> '';                     {Loop while not COM reference}
  Result := 'COM' + Result;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.ExtractPortDescription(Port: String): String;
{Returns Port Description (everything except (COMxxx)) from Port string.}
var
  PIdx   : Integer;
  PortID : String;
begin
  {Start with full description, then find PortID portion}
  Result := Port;
  PortID := ExtractPortID(Result);
  PIdx := pos(PortID, uppercase(Result));
  if PIdx > 0 then
    begin {Strip out PortID and surrounding brackets, if any}
    if (PIdx + length(PortID) <= length(Result)) and (Result[PIdx + length(PortID)] in [')', ']', '}']) then PortID := PortID + Result[PIdx + length(PortID)];
    if (PIdx > 1) and (Result[PIdx-1] in ['(', '[', '{']) then PortID := Result[PIdx-1]+PortID;
    delete(Result, pos(PortID, uppercase(Result)), length(PortID));
    end;
  {Trim right side of control characters and white space}
  while (length(Result) > 0) and (Result[length(Result)] in [chr(0)..chr(32), chr(128)..chr(255)]) do delete(Result, length(Result), 1);
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.FilterPorts;
{Filter Ports by Rules (sorting and filtering).  Results returned in FPorts[], FPortIDs[], FPortDescs[], and FProperties[].}
var
  PortIdx    : Integer;
  RuleIdx    : Integer;
  IncAll     : Boolean;
  CPorts     : TStrings;

const
  States : array['+'..'9'] of integer = (1, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

    {----------------}

    procedure UpdatePort(PortID: String; State, RuleIndex: Integer);
    {Add or Update port details.  PortID is used to identify the port to added to the filtered dataset.
     State is -1, 0, or 1, indicating desired exclude, keep, include state.  RuleIndex specifies the index of the rule that is being applied.}
    var
      PIdx1, PIdx2           : Integer;
      Properties             : PPortProperties;

        {----------------}

        function GetState(Present: Boolean): TPortState;
        {Returns proper state, based on Present (or absent) status, IncAll flag, and desired State}
        begin
          case State of
            -1 : Result := TPortState(ifthen(Present, ord(psPresentExcluded), ord(psAbsentExcluded))); {Exclude}
             0 : Result := TPortState(ifthen(Present, ifthen(IncAll, ord(psPresentIncluded), ord(psPresentExcluded)), ifthen(IncAll, ord(psAbsentIncluded), ord(psAbsentExcluded)))); {Current}
             1 : Result := TPortState(ifthen(Present, ord(psPresentIncluded), ord(psAbsentIncluded))); {Include}
          end;
        end;

        {----------------}

    begin
      {Search for PortID in current (present) ports list}
      PIdx1 := 0;
      while (PIdx1 < CPorts.Count) and (ExtractPortID(CPorts[PIdx1]) <> PortID) do inc(PIdx1);
      {Search for PortID in filtered port list}
      PIdx2 := FPortIDs.IndexOf(PortID);
      if PIdx2 > -1 then
        begin    {Filtered port found}
        if PIdx1 < CPorts.Count then
          CPorts.Delete(PIdx1)          {Duplicate Port ID, delete it}
        else                            {Else, update unique port properties}
          if (State <> 0) and (not PPortProperties(FProperties[PIdx2])^.InclExcl) then
            begin
            PPortProperties(FProperties[PIdx2])^.State := GetState(PPortProperties(FProperties[PIdx2])^.State in [psPresentIncluded, psPresentExcluded]);
            PPortProperties(FProperties[PIdx2])^.InclExcl := True;
            PPortProperties(FProperties[PIdx2])^.InclExclIdx := RuleIndex;
            end;
        end
      else
        begin    {Port not yet in filtered list}
        if PIdx1 < CPorts.Count then
          begin  {Present port found}
          FPorts.Add(CPorts[PIdx1]);
          FPortDescs.Add(ExtractPortDescription(CPorts[PIdx1]));
          getmem(Properties, sizeof(TPortProperties));
          Properties.State := GetState(True);
          CPorts.Delete(PIdx1);
          end
        else
          begin  {Port not found}
          {Exit if rule not manually including or excluding port}
          if State = 0 then exit;
          {Otherwise, add it}
          FPorts.Add('<port unknown> ('+PortID+')');
          FPortDescs.Add('<port unknown>');
          getmem(Properties, sizeof(TPortProperties));
          Properties.State := GetState(False);
          end;
        Properties.InclExcl := State <> 0;
        Properties.InclExclIdx := ifthen(Properties.InclExcl, RuleIndex, -1);
        FPortIDs.Add(PortID);
        FProperties.Add(Properties);
        end;
    end;

    {----------------}

begin
  CPorts := TStringList.Create;
  try
    {Clear current port metrics}
    FPorts.Clear;
    FPortIDs.Clear;
    FPortDescs.Clear;
    while (FProperties.Count > 0) do
      begin
      freemem(PPortProperties(FProperties[0]), sizeof(TPortProperties));
      FProperties.Delete(0);
      end;
    FPortsExcluded := 0;

    {Set global rule to default (include all)}
    IncAll := True;

    {Get current O.S.-specified ports}
    CPorts.Assign(FOSPorts);

    if FFiltered then
      begin {If allowed to filter ports by rules...}
      {Apply Global Rule (include all or exclude all)}
      if FRules.Count > 0 then IncAll := GlobalInclude;

      {Apply Sort-Order Port ID rules (may also be a combined sort-order and include/exclude rule, both properties of which are applied here)}
      for RuleIdx := 1 to FRules.Count-1 do
        if (FRules[RuleIdx][1] = '(') then {Sort-Order Port ID Rule}
          UpdatePort('COM'+inttostr(abs(strtoint(FRules[RuleIdx]))), States[FRules[RuleIdx][2]], RuleIdx);
      end;

    {Append any remaining Ports not matching sort-order rules}
    while CPorts.Count > 0 do UpdatePort(ExtractPortID(CPorts[0]), 0, -1);

    if FFiltered then
      begin {If allowed to filter ports by rules...}
      {Apply non-Sort-Order, Include/Exclude Port ID rules}
      for RuleIdx := 1 to FRules.Count-1 do
        if (FRules[RuleIdx][1] in ['-','+']) and (Frules[RuleIdx][2] in ['0'..'9']) then {Non-Sort-Order Port ID Rule}
          UpdatePort('COM'+inttostr(abs(strtoint(FRules[RuleIdx]))), States[FRules[RuleIdx][1]], RuleIdx);

      {Apply Include/Exclude Port Description rules}
      for RuleIdx := 1 to FRules.Count-1 do
        begin
        if (FRules[RuleIdx][1] in ['-','+']) and not (FRules[RuleIdx][2] in ['0'..'9']) then {Include/Exclude Port Description Rule}
          begin
          PortIdx := -1;
          repeat
            PortIdx := IndexOfPortDesc(rightstr(FRules[RuleIdx], length(FRules[RuleIdx])-1), PortIdx);
            if PortIdx > -1 then UpdatePort(FPortIDs[PortIdx], States[FRules[RuleIdx][1]], RuleIdx);
          until PortIdx = -1;
          end;
        end;

      {Need Scannable ports only?}
      if FScannable then
        begin {If port list must contain only scannable ports...}
        PortIdx := 0;
        while (PortIdx < FPorts.Count) do
          if not (PPortProperties(FProperties[PortIdx])^.State = psPresentIncluded) then
            begin {Port not currently scannable; remove it}
            inc(FPortsExcluded, ord(not (PPortProperties(FProperties[PortIdx])^.State in [psAbsentIncluded, psAbsentExcluded])));
            FPorts.Delete(PortIdx);
            FPortIDs.Delete(PortIdx);
            FPortDescs.Delete(PortIdx);
            freemem(PPortProperties(FProperties[PortIdx]), sizeof(TPortProperties));
            FProperties.Delete(PortIdx);
            end
          else
            inc(PortIdx);
        end;
      end;

  finally
    CPorts.Free;
  end;
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.ParseSearchRuleString(RuleStr: String);
{Parse and extract serial port search rules from RuleStr into FRules.
 Format: FRules is filled in the following way:
         Element  0:   (Global Rule)  -  May be +* or -*.  +* indicates to include all ports unless excluded by another rule.  -* indicates to exclude all ports unless included by another rule.
         Elements 1-n: (Detail Rules) -  There are two classes of detail rules: 1) Port ID, and 2) Port Description.
                       Port ID:          May be -#, +#, (#), (-#), or (+#) ;where # is one or more numeric digits representing a serial port ID.
                                         -# indicates to exclude a port.  +# indicates to include a port.  A Port ID rule in parentheses indicates its search order; its order of appearance in
                                         relation to other parenthesized Port ID rules.  (#) indicates a port's search order if it is already present; ie: not manually included or excluded.
                       Port Description: May be -x, or +x ; where x is one or more characters (beginning with a non-number character) representing a port description string to match.
                                         -x indicates to exclude a port if its description matches the string.  +x indicates to include a port if its description matches the string.
                                         The order of Port Description rules in elements 1-n is inconsequential, thus, parenthesized Port Description rules are not allowed.}
var
  RuleIdx : Integer;
  R       : String;
const
  IsIn      = True;
  IsNotIn   = False;
  StripChar = True;
  LeaveChar = False;

    {----------------}

    function NextRule(var Rule: String): Boolean;
    {Return next Rule from RuleStr as well as True if rule found.}
    var
      CIdx  : Integer;
      Order : Boolean;


        {----------------}

        function NextChar(Includes: Boolean; Chrs: String; DelChar: Boolean = False; CPos: Integer = 1): Boolean;
        {Strips leading spaces at CPos, then:
           if Includes = True : returns True if next character is one of Chrs, False otherwise.
           if Includes = False: returns True if next character is not any one of Chrs, False otherwise.
         If Result = True and DelChar = True, deletes the char after comparison.}
        var
          Idx : Cardinal;
        begin
          Result := not Includes;
          while (Rule <> '') and (Rule[CPos] = ' ') do delete(Rule, CPos, 1);                  {Strip leading spaces}
          if length(Rule) < CPos then exit;
          if Includes then
            for Idx := 1 to length(Chrs) do Result := Result or (Rule[CPos] = Chrs[Idx])
          else
            for Idx := 1 to length(Chrs) do Result := Result and (Rule[CPos] <> Chrs[Idx]);
          delete(Rule, CPos, ord(Result and DelChar));
        end;

        {----------------}

    begin
      Rule := '';
      while (Rule = '') and (RuleIdx <= length(RuleStr)) do
        begin {While rule not found but more to parse}
        try
          while (RuleIdx <= length(RuleStr)) and ((RuleStr[RuleIdx] <> ',') or ((RuleIdx > 1) and (RuleStr[RuleIdx-1] = '\'))) do
            begin                                                                              {Retrieve next rule (delimited by comma ',' or end of string)}
            Rule := Rule + RuleStr[RuleIdx];
            inc(RuleIdx);
            end;
          inc(RuleIdx);
          Order := NextChar(IsIn, '(', StripChar);                                             {Is possible sort-order rule?}
          if Order then
            begin                                                                              {Sort-order rule}
            if NextChar(IsNotIn, '-+0123456789') then Abort;                                   {  If invalid sort-order rule, skip to next rule}
            if NextChar(IsIn, '-+') and NextChar(IsNotIn, '0123456789', LeaveChar, 2) then Abort;
            end
          else                                                                                 {Non-sort-order rule}
            if NextChar(IsNotIn, '-+') or (length(Rule) < 2) then Abort;                       {  If invalid non-sort-order rule, skip to next rule}
          if Order or (NextChar(IsIn, '-+') and NextChar(IsIn, '0123456789', False, 2)) then   {If Port ID rule}
            begin                                                                              {  Strip any trailing non-numbers}
            CIdx := 2;
            while (CIdx <= length(Rule)) and (Rule[CIdx] in ['0'..'9']) do inc(CIdx);
            Rule := leftstr(Rule, CIdx-1);
            if Order then Rule := '(' + leftstr(Rule, CIdx-1) + ')';                           {If sort-order rule, reformat as such}
            end;
        except
          Rule := '';                                                                          {Exception?  Must be invalid rule, clear and try again}
        end;
        end;  {While rule not found}
      Result := Rule <> '';
    end;

    {----------------}

    procedure PlaceRule(Rule: String);
    {Place Rule into FRules, replacing and combining rules as necessary}
    var
      Idx      : Integer;
      TargetID : Integer;
    begin
      if (Rule = IncludeAll) or (Rule = ExcludeAll) then                                       {Global rule?}
        FRules[0] := Rule                                                                      {  Replace existing rule}
      else
        if (Rule[1] in ['-','+']) and not (Rule[2] in ['0'..'9']) then                         {Else, Port Description rule?}
          FRules.Add(Rule)                                                                     {  Add rule}
        else                                                                                   {Else, Port ID rule}
          begin                                                                                {  Add or combine with existing rule}
          Idx := 0;
          TargetID := abs(strtoint(Rule));
          while (Idx < FRules.Count) and (abs(strtoint(FRules[Idx])) <> TargetID) do inc(Idx);
          if Idx >= FRules.Count then
            FRules.Add(Rule)
          else
            FRules[Idx] := ifthen((FRules[Idx][1] = '(') and (Rule[1] <> '('), '(' + Rule + ')', Rule);
          end;
    end;

    {----------------}

begin
  TStringList(FRules).OnChange := nil;                                                       {Disable onchange event}
  FRules.Clear;
  FRules.Add(IncludeAll);                                                                    {Include all ports by default}
  RuleIdx := 1;
  while NextRule(R) do PlaceRule(R);                                                         {Parse and place all rules}
  TStringList(FRules).OnChange := RulesChanged;                                              {Re-enable onchange event}
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GenerateSearchRuleString: String;
{Return search rule string representing the search rules contained in FRules.
 Format: All rules are comma (,) separated.
         The first rule is the global include/exclude rule (+* or -*).
         The following rules are Port ID rules; with parenthesized rules listed first.
         The remaining rules are Port Description rules.}
var
  Idx : Integer;
begin
  {Begin with global include/exclude rule (or default string)}
  Result := ifthen(FRules.Count > 0, FRules[0], IncludeAll);
  {Follow with all sort-order Port ID rules}
  for Idx := 1 to FRules.Count-1 do if FRules[Idx][1] = '(' then Result := Result + ',' + FRules[Idx];
  {Then with all non-sort-order Port ID rules}
  for Idx := 1 to FRules.Count-1 do if (FRules[Idx][1] <> '(') and (FRules[Idx][2] in ['0'..'9']) then Result := Result + ',' + FRules[Idx];
  {End with all Port Description rules}
  for Idx := 1 to FRules.Count-1 do if (FRules[Idx][1] <> '(') and not (FRules[Idx][2] in ['0'..'9']) then Result := Result + ',' + FRules[Idx];
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.RulesChanged(Sender: TObject);
{FRules changed externally; cleans by cycling through RuleString then back again.}
begin
  ParseSearchRuleString(GenerateSearchRuleString);
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetCount: Integer;
{Return count of elements.  This count is affected by the state of FFiltered and FScannable.}
begin
  Result := FPorts.Count;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetGlobal(Index: Integer): Boolean;
{Returns state of global rule.}
begin
  Result := True; {Default is to include all}
  case Index of
    0 : Result := (FRules.Count > 0) and (FRules[0][1] = '+'); {Global Include All}
    1 : Result := (FRules.Count > 0) and (FRules[0][1] = '-'); {Global Exclude All}
  end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetPort(Index, DataType: Integer): String;
{Return Port string (full description, ID, or description).  This result is affected by the state of FFiltered and FScannable.
 DataType indicates what type of data to retrieve.  Index is the element to retrieve from.}
begin
  Result := '';
  if Index >= FPorts.Count then exit;   {Exit if out of range. (Note FPort should be same count as FPortIDs and FPortDescs)}
  case DataType of
    0: Result := FPorts[Index];      {Get entire port description (including Port ID)}
    1: Result := FPortIDs[Index];    {Get Port ID}
    2: Result := FPortDescs[Index];  {Get Port Description (without Port ID)}
  end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetProperty(Index, DataType: Integer): Boolean;
{Return property of port.  This result is affected by the state of FFiltered and FScannable.
 DataType indicates what type of data to retrieve.  Index is the element to retrieve from.}
begin
  Result := False;
  if Index >= FProperties.Count then exit;   {Exit if out of range.}
  case DataType of
    0: Result := PPortProperties(FProperties[Index])^.State in [psPresentIncluded, psPresentExcluded];    {Get Present status of port}
    1: Result := PPortProperties(FProperties[Index])^.State in [psAbsentIncluded, psPresentIncluded];     {Get Included status of port}
    2: Result := PPortProperties(FProperties[Index])^.State in [psAbsentExcluded, psPresentExcluded];     {Get Excluded status of port}
    3: Result := PPortProperties(FProperties[Index])^.InclExcl;                                           {Get InclExcl-by-rule status of port}
  end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetDescRulesString: String;
{Return string of current Port Description rules only (no sort rules or Port ID rules)}
var
  Idx : Integer;
begin
  Result := '';
  for Idx := 1 to FRules.Count-1 do
    if (FRules[Idx][1] in ['-','+']) and not (FRules[Idx][2] in ['0'..'9']) then Result := Result + ',' + FRules[Idx];
  if Result <> '' then delete(Result, 1, 1);
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetIDRulesString: String;
{Return string of current include/exclude Port ID rules only (will include sort rules if they are include/exclude rules)}
var
  Idx : Integer;
begin
  Result := '';
  for Idx := 1 to FRules.Count-1 do
    if ((FRules[Idx][1] = '(') and (FRules[Idx][2] in ['-','+'])) or
       ((FRules[Idx][1] in ['-','+']) and (FRules[Idx][2] in ['0'..'9'])) then Result := Result + ',' + FRules[Idx];
  if Result <> '' then delete(Result, 1, 1);
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetSortRulesString: String;
{Return string of current sort rules only (no inclusion/exclusion rules)}
var
  Idx : Integer;
begin
  Result := '';
  for Idx := 1 to FRules.Count-1 do
    if FRules[Idx][1] = '(' then Result := Result + ',(' + inttostr(abs(strtoint(FRules[Idx]))) + ')';
  if Result <> '' then delete(Result, 1, 1);
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetRuleIdx(Index: Integer): Integer;
{Returns index of include/exclude rule that was applied to filtered port at Index.
 Returns -1 if no rule applied.}
begin
  Result := -1;
  if (Index < 0) or (Index >= FPorts.Count) then exit;
  if PPortProperties(FProperties[Index])^.InclExcl then Result := PPortProperties(FProperties[Index])^.InclExclIdx;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.GetRuleType(Index, DataType: Integer): Boolean;
{Returns True or False if the rule at Index matches DataType}
begin
  Result := False;
  if (Index < 0) or (Index >= FRules.Count) then exit;
  case DataType of
    0: Result := FRules[Index][1] = '(';                                                   {Is Sort-Order Port ID rule? (may also be an Include/Exclude Port ID rule)}
    1: Result := ((FRules[Index][1] = '(') and (FRules[Index][2] in ['-','+'])) or         {Is Include/Exclude Port ID rule?}
                 ((FRules[Index][1] in ['-','+']) and (FRules[Index][2] in ['0'..'9']));
    2: Result := (FRules[Index][1] in ['-','+']) and not (FRules[Index][2] in ['0'..'9']); {Is Port Description rule?}
  end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.IndexOfPortDesc(MatchDesc: String; AfterIdx: Integer = -1): Integer;
{Return index of port matching MatchDesc, starting with ports AfterIdx.  MatchDesc may contain asterisks (*) as a wildcard.
 Returns -1 if none found.}
var
  MatchParts : TStrings;
  Beginning  : Boolean;
  Ending     : Boolean;

        {----------------}

        function WildcardPos(MDesc: String): Integer;
        {Return first character position of MDesc containing the wildcard character (*).  Note escaped asterisks (\*) are not wildcard characters.}
        var
          Idx : Integer;
          Esc : Boolean;
          Len : Integer;
        begin
          Idx := 1;
          Esc := False;
          Len := length(MDesc);
          while (Idx <= Len) and ((MDesc[Idx] <> '*') or Esc) do
            begin {while not end of string and character is not a wildcard (*) or character is an escaped asterisk (\*)...}
            Esc := not Esc and (MDesc[Idx] = '\');
            inc(Idx);
            end;
          Result := ifthen(Idx <= Len, Idx, 0);
        end;

        {----------------}

        procedure GetMatchParts;
        {Split MatchDesc into MatchParts and set Beginning and Ending flags to indicate whether or not MatchDesc requires matching start and/or end.}
        var
          CIdx : Integer;
        begin
          Beginning := (length(MatchDesc) > 0) and (MatchDesc[1] <> '*');
          Ending := (length(MatchDesc) > 0) and ((MatchDesc[length(MatchDesc)] <> '*') or ((length(MatchDesc) > 1) and (MatchDesc[length(MatchDesc)-1] = '\')));
          repeat
            CIdx := WildcardPos(MatchDesc);
            if CIdx > 0 then
              begin {Wildcard found}
              if CIdx > 1 then MatchParts.Add(UnEscapeDesc(leftstr(MatchDesc, CIdx-1)));
              delete(MatchDesc, 1, CIdx);
              end
            else
              begin {Wildcard not found}
              MatchParts.Add(UnEscapeDesc(MatchDesc));
              MatchDesc := '';
              end;
          until MatchDesc = '';
        end;

        {----------------}

        function SearchSubPart(MPIdx: Integer; PortDesc: String): Boolean;
        {Recursively search PortDesc for each of MatchParts.  MPIdx must be 0 from main caller.
         Returns True if PortDesc is a match for MatchParts, False otherwise.}
        var
          CIdx   : Integer;
        begin
          Result := True;
          if MPIdx = MatchParts.Count then exit;
          CIdx := pos(uppercase(MatchParts[MPIdx]), uppercase(PortDesc));
          Result := not ((CIdx < 1) or ((MPIdx = 0) and Beginning and (CIdx > 1)) or ((MPIdx = MatchParts.Count-1) and Ending and (CIdx < length(PortDesc)-length(MatchParts[MatchParts.Count-1])+1)));
          if Result and (MPIdx < MatchParts.Count-1) then {Sub part matched and more parts to go, continue search}
              Result := Result and SearchSubPart(MPIdx+1, rightstr(PortDesc, length(PortDesc)-(CIdx+length(MatchParts[MPIdx])-1)));
          if not Result and (CIdx > 0) then Result := SearchSubPart(MPIdx, rightstr(PortDesc, length(PortDesc)-CIdx));
        end;

        {----------------}

begin
  MatchParts := TStringList.Create;
  try
    GetMatchParts;
    Result := AfterIdx;
    repeat inc(Result); until (Result >= FPorts.Count) or SearchSubPart(0, FPortDescs[Result]);
    if Result >= FPorts.Count then Result := -1;
  finally
    MatchParts.Free;
  end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.RuleMatchCount(MatchDesc: String): Integer;
{Returns number of ports matching MatchDesc}
var
  Idx : Integer;
begin
  Result := 0;
  if (length(MatchDesc) < 2) or (NeedEscapeDesc(MatchDesc) > 0) then exit;
  if MatchDesc[1] in ['-','+'] then delete(MatchDesc, 1, 1);
  Idx := -1;
  repeat
    Idx := IndexOfPortDesc(MatchDesc, Idx);
    inc(Result, ord(Idx > -1));
  until Idx = -1;
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.SetFiltered(Value: Boolean);
{Set FFiltered field and re-filter Ports}
begin
  FFiltered := Value;
  FilterPorts;
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.SetScannable(Value: Boolean);
{Set FScannable field and re-filter Ports}
begin
  FScannable := Value;
  FilterPorts;
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.SetFOnReadPortRules(Value: TReadPortRulesEvent);
{Set FOnReadPortRules event handler field and force trigger of event}
begin
  FOnReadPortRules := Value;
  DoReadPortRules;
end;

{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบ     Public Routines     บบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}

constructor TPortMetrics.Create;
begin
  inherited Create;
  FOSPorts    := TStringList.Create;
  FPorts      := TStringList.Create;
  FPortIDs    := TStringList.Create;
  FPortDescs  := TStringList.Create;
  FProperties := TList.Create;
  FRules      := TStringList.Create;
  TStringList(FRules).CaseSensitive := False;                                                {Set to ignore case}
  FRules.Add(IncludeAll);                                                                    {Include all ports by default}
  TStringList(FRules).OnChange := RulesChanged;                                              {Attached FRules' change event to handler}
  FFiltered   := True;
  FScannable  := False;
  FOnDeviceChange := nil;
  FOnReadPortRules := nil;
  FOnWritePortRules := nil;
  FPortsExcluded := 0;
  Refresh;
end;

{------------------------------------------------------------------------------}

destructor TPortMetrics.Destroy;
begin
  FRules.Free;
  while (FProperties.Count > 0) do
    begin
    freemem(PPortProperties(FProperties[0]));
    FProperties.Delete(0);
    end;
  FProperties.Free;
  FPortDescs.Free;
  FPortIDs.Free;
  FPorts.Free;
  FOSPorts.Free;
  inherited Destroy;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.EscapeDesc(Description: String): String;
{Returns Description string with properly "escaped" special characters, if they appear}
var
  Idx       : Integer;
begin
  Idx := 1;
  Result := Description;
  while Idx <= length(Result) do
    begin
    if (Result[Idx] in ['*', ',', '\']) then
      begin {Found special char, prepend escape}
      insert('\', Result, Idx);
      inc(Idx, 2);
      end
    else    {Found normal character}
      inc(Idx);
    end;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.UnEscapeDesc(Description: String): String;
{Returns Description string without "escaped" special characters, if they appear}
var
  Idx : Integer;
begin
  Result := Description;
  for Idx := 1 to length(Result) do if (Result[Idx] = '\') then delete(Result, Idx, 1);
end;

{------------------------------------------------------------------------------}

function TPortMetrics.NeedEscapeDesc(Description: String): Integer;
{Returns index of character in Description that needs escaping; 0 if none}
var
  Esc : Boolean;
begin
  Result := 1;
  Esc := False;
  while (Result <= length(Description)) and (not (Description[Result] in [',', '\']) or Esc or ((Description[Result] = '\') and ((Result < length(Description)) and (Description[Result+1] in ['\',',','*'])))) do
    begin
    Esc := not Esc and (Description[Result] = '\');
    inc(Result);
    end;
  if Result > length(Description) then Result := 0;
end;

{------------------------------------------------------------------------------}

function TPortMetrics.IndexOfPortID(PortID: String): Integer;
{Return index of PortID, if exists.  Returns -1 if not found.}
begin
  Result := FPortIDs.IndexOf(PortID);
end;

{------------------------------------------------------------------------------}

procedure TPortMetrics.Refresh;
{Update entire object (dataset) with port list from O.S. and filtered against FRules (if necessary).}
begin
  EnumerateComPorts(True, FOSPorts);
  FilterPorts;
end;

{##############################################################################}
{##############################################################################}
{########################### TPHintWindow Routines ############################}
{##############################################################################}
{##############################################################################}

{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบ     Private Routines     บบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}

function TPHintWindow.CanHint(var NewText: String): Boolean;
{Perform OnCanHint event and return True if we can display hint, False otherwise.
 NewText is filled with the new text to display.}
begin
  Result := True;
  if assigned(FCanHint) then
    begin
    NewText := Text;
    Result := FCanHint(FControl, FItemID, NewText);
    end
  else
    NewText := FControl.Hint;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.DoShowHint(NewInterval: Integer = 0);
{Show the hint.  Next interval is set to FHidePause unless NewInterval is provided.}
var
  HintRect : TRect;

    {------------------------}

    function Centered(Str: String): String;
    {Center Str (requires CR/LF combinations in string)}
    var
      Idx, LineWidth, CenterPos, SpaceWidth : Integer;
      LineStr                               : String;
    begin
      Result := '';
      CenterPos := (HintRect.Right-HintRect.Left) div 2 - 3;
      SpaceWidth := Canvas.TextWidth(stringofchar(' ', 1));
      while Str <> '' do
        begin
        Idx := pos(#$D#$A, Str);
        LineStr := ifthen(Idx > 0, leftstr(Str, Idx-1), Str);
        delete(Str, 1, ifthen(Idx > 0, Idx+1, length(Str)));
        LineWidth := Canvas.TextWidth(LineStr);
        Result := Result + stringofchar(' ', (CenterPos-(LineWidth div 2)) div SpaceWidth)+LineStr+ifthen(Str <> '', #$D#$A, '');
        end;
      Result := Result + Str;
    end;

    {------------------------}

begin
  {Configure Hint}
  FTimer.Enabled := False;
  FTimer.Interval := ifthen(NewInterval > 0, NewInterval, FHidePause);
  FTimer.OnTimer := TriggerHideHint;
  FTimer.Enabled := True;
  HintRect := CalcHintRect(FMaxWidth, Text, nil);
  {Position and display hint}
  HintRect.Left := HintRect.Left + FPos.X;
  HintRect.Top := HintRect.Top + FPos.Y;
  HintRect.Right := HintRect.Right + FPos.X;
  HintRect.Bottom := HintRect.Bottom + FPos.Y;
  if not FCentered then ActivateHint(HintRect, Text) else ActivateHint(HintRect, Centered(Text));
  FShowing := True;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.DoHideHint(NewInterval: Integer = 0);
{Hide hint window for port row.  Next interval is set to FHintPause unless NewInterval is provided.}
begin
  FTimer.Enabled := False;
  FTimer.Interval := ifthen(NewInterval > 0, NewInterval, FShowPause);
  FTimer.OnTimer := TriggerShowHint;
  ReleaseHandle;
  FShowing := False;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.SetControl(Value: TControl);
{Set parent control of this hint object}
begin
  DisableHint(True);
  FControl := Value;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.SetNoItemID(Value: Integer);
{Set NoItemID value.}
begin
  if FItemID = FNoItemID then FItemID := Value;
  FNoItemID := Value;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.SetPos(Value: TPoint);
{Set position of upper-left corner of hint window}
begin
  FPos := Value;
  if FShowing then DoShowHint;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.TriggerShowHint(Sender: TObject);
{Call DoShowHint.  This method is called by FTimer's OnTimer event.}
begin
  DoShowHint;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.TriggerHideHint(Sender: TObject);
{Call DoHideHint.  This method is called by FTimer's OnTimer event.}
begin
  DoHideHint;
end;

{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบ     Public Routines     บบบบบบบบบบบบบบบบบบบบบบบบบบ}
{บบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบบ}

constructor TPHintWindow.Create(Control: TControl);
begin
  inherited Create(Control);
  FControl := Control;
  FCentered := False;
  FNoItemID := MAXINT;
  FItemID := FNoItemID;
  FShowing := False;
  FCanHint := nil;
  FHidePause := Application.HintHidePause;
  FShowPause := Application.HintPause;
  FShortPause := Application.HintShortPause;
  FMaxWidth := 500;
  Color := Application.HintColor;
  FTimer := TTimer.Create(FControl);
  FTimer.Enabled := False;
  FTimer.OnTimer := TriggerShowHint;
  FTimer.Interval := FShowPause;
end;

{------------------------------------------------------------------------------}

destructor TPHintWindow.Destroy;
begin
  inherited Destroy;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.DisableHint(ClearID: Boolean);
{Disable the hint window.  If ClearID = True, FItemID set back to NoItemID.}
begin
  if GetCaptureControl = FControl then SetCaptureControl(nil);
  DoHideHint;
  if ClearID then FItemID := FNoItemID;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.HideHint;
{Hide hint window.}
begin
  DoHideHint;
end;

{------------------------------------------------------------------------------}

procedure TPHintWindow.UpdateHintMetrics(ItemID: Integer; X, Y: Integer);
{Update hint metrics and enable, disable, or re-enable hint as necessary}
var
  Pos     : TPoint;
  NewText : String;
begin
  if (X < 0) or (X > FControl.Width) or (Y < 0) or (Y > FControl.Height) then ItemID := FNoItemID;          {Check for out-of-bounds condition}
  if ItemID = FNoItemID then                                                                                {Mouse moved outside of items' boundaries?}
    DisableHint(True)                                                                                       {  Disable hint}
  else                                                                                                      {Else, mouse within item boundary}
    begin
    if GetCaptureControl = nil then SetCaptureControl(FControl);                                            {  Capture mouse if we haven't already}
    Pos := FControl.ClientToScreen(point(X, Y+20));                                                         {  Convert coordinates to screen coordinates}
    if (FPos.X = Pos.X) and (FPos.Y = Pos.Y) then exit;                                                     {  Mouse position unchanged?  Exit}
    if not FShowing then FPos := Pos;                                                                       {  If hint not showing yet, update position}
    if ItemID <> FItemID then                                                                               {  Mouse over new item?}
      begin
      FItemID := ItemID;                                                                                    {    Remember port's row index}
      if CanHint(NewText) then                                                                              {    Item should display hint?}
        begin
        DoHideHint(FShortPause*ord(FShowing));                                                              {      Hide/disable current hint,}
        if NewText <> Text then Text := NewText;                                                            {      update hint text, if necessary, and}
        FTimer.Enabled := True;                                                                             {      re-enable hint (with immediacy if}
        end                                                                                                 {      another hint was already showing)}
      else                                                                                                  {    Else, item shouldn't be hinted}
        DoHideHint;                                                                                         {      Hide hint}
      end;
    end;
end;

{------------------------------------------------------------------------------}

initialization
  COM := TPortMetrics.Create;
  SelRow := 0;
  OldSelRow := 0;
  UndoRules := '';
  RedoRules := '';
  {Allocate memory for Window Placement structure and set its length}
  GetMem(WinPos,sizeof(TWindowPlacement));
  WinPos.length := sizeof(TWindowPlacement);

finalization
  COM.Destroy;
  {Free WinPos memory}
  freemem(WinPos);

end.
