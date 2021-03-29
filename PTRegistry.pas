unit PTRegistry;

interface
uses
  Windows, WinTypes, Classes, SysUtils, StrUtils, Registry, Math, Forms, Dialogs, Controls, ShlObj, MultiMon, PortList;

const
  CustomSchemeID = 3;

type
  {Define Preference Entities}
  {Preference settings (entities) may be entered here in any order.  To add a new one, first add it to this enumeration, then create
  its default in the DPrefs array.

  To include a preference entity in the Preferences form, if the setting is related directly to a control type (checkbox, radio group, etc), append the
  new entity to the PrefTx array (in Prefs unit) and add the appropriate control to the Preferences form with a tag equal to its entity's place in the
  PrefTx array.  Otherwise, additional code to support the related control must be added to the Prefs unit.}
  PrefEntity = (FontSize, ChartFontSize, ChartDisplayMode, InfoShowHex, CONTabs, VARTabs, OBJTabs, PUBPRITabs, DATTabs, ExplorerVisible, EditorPos,
                ExplorerWidth, ExplorerPanelSplitPos, FileSplitPos, FindReplacePos, CharChartPos, PrefsPos, PortListPos, LibraryPaths, AutoRecover,
                UndoAfterSave, File01, File02, File03, File04, File05, File06, File07, File08, File09, File10, CheckAssoc, ExtSpin, ExtSpin2, ExtBinary,
                ExtEEPROM, ExtFlash, Find01, Find02, Find03, Find04, Find05, Find06, Find07, Find08, Find09, Find10, Replace01, Replace02, Replace03,
                Replace04, Replace05, Replace06, Replace07, Replace08, Replace09, Replace10, TopFile, ShowBookmarks, ShowBlockIndentions, ShowLineNumbers,

                {AutoIndent, BackspaceAutoUnindents, HighlightSyntax, FavoriteDirectory01,
                FavoriteDirectory02, FavoriteDirectory03, FavoriteDirectory04, FavoriteDirectory05, FavoriteDirectory06,
                FavoriteDirectory07, FavoriteDirectory08, FavoriteDirectory09, FavoriteDirectory10, StartupDirectory}

                LastUsedDirectory, ShowRecentOnly, FilterIdx, SingleInstanceOnly, ResetSignal, ResetDelay, ResetDelayP2, SerialSearch, SerialSearchRules,
                Initialized,

                SynScheme, SynSchemePath, SynSchemeModified,

                {NOTE: If any of the following "Syn" preferences change in any way, update FirstSynPref and LastSynPref constants, SynSchemes (GlobalSecondary),
                and ElementInfo (Prefs) appropriately.}
                SynRegular, SynBlocksCONstant, SynBlocksVARiable, SynBlocksOBJect, SynBlocksPUBlic, SynBlocksPRIvate, SynBlocksDATa,
                SynSpinDebug, SynSpinIADebug, SynSpinCommands, SynSpinConditionals, SynSpinVariables,
                SynAssemblyDebug, SynAssemblyInstructions, SynAssemblyConditionals, SynAssemblyEffects, SynAssemblyLiterals,
                SynDirectives, SynCommentsCode, SynCommentsDocument, SynConstantsNumbers, SynConstantsPredefined, SynStrings, SynOperators, SynRegisters,
                SynSizes, SynDataAlignments, SynDataSizes,

                NewP1FileTemplate, NewP2FileTemplate);

  {Define editor preferences record type}
  PrefType   = (StringType, IntegerType, BooleanType); {The enumerated type of preference}
  TPrefs = record
    Name          : ShortString;                       {The name of the registry key}
    PType         : PrefType;                          {The Preference Type}
    BValue        : Boolean;                           {The Boolean value (PType = BooleanType)}
    IValue        : Cardinal;                          {The Integer value (PType = IntegerType)}
    SValue        : String;                            {The String value (PType = StringType)}
  end;

  procedure OpenRegistry(TheKey: HKEY);
  procedure CloseRegistry;
  function  UpgradePrefsFromRegistry: Boolean;
  function  LoadPrefsFromRegistry: Boolean;
  procedure PerformFinalPrefLoading;
  procedure SavePrefsToRegistry;
  function  VersionInitialized: Boolean;  
  procedure CheckFileAssociationsInRegistry(ForceAssociations: Boolean);
  //procedure ReadPrefs(Parent: TWinControl; Prefs: array of TPrefs);
  //procedure WritePrefs(Parent: TWinControl);
  function  GetPosValue(PosStr: String; Position: Integer): Integer;
  procedure LoadWindowMetrics(Window: TForm; Pref: PrefEntity);
  procedure SaveWindowMetrics(Window: TForm; Pref: PrefEntity);
  function  GetDisplayableWindowPosition(Bounds: TRect): TRect;
  procedure EnsureWindowDisplayable(Window: TForm);

const
  {Starting and ending syntax preference entities}
  FirstSynPref = SynRegular;
  LastSynPref = SynDataSizes;

var
  CPrefs           : array[PrefEntity] of TPrefs;                          {The current Editor Preferences}
  Prefs            : TRegistry;
  PrefsHaveChanged : Boolean;
  PIdx             : PrefEntity;
  VerInitializing  : Boolean;                                              {True = current session is the first time this version has run}
  WinPos           : PWindowPlacement;                                     {Structure to describe window position status}

const
  EditorRegistryVersion = '2.4.0';       {The registry version used by this editor; should only change with a legacy-incompatible registry issue}
  RegPath = '\SOFTWARE\ParallaxInc\Propeller\';

  DPrefs: array [PrefEntity] of TPrefs =  {The default preferences settings}
    ( (NAME: 'FontSize';                 PTYPE: IntegerType;   IVALUE: 12),
      (NAME: 'ChartFontSize';            PTYPE: IntegerType;   IVALUE: 36),
      (NAME: 'ChartDisplayMode';         PTYPE: IntegerType;   IVALUE: 0),
      (NAME: 'InfoShowHex';              PTYPE: BooleanType;   BVALUE: False),
      (NAME: 'CONTabs';                  PTYPE: StringType;    SVALUE: '2,8,16,18,32,56,80'),
      (NAME: 'VARTabs';                  PTYPE: StringType;    SVALUE: '2,8,22,32,56,80'),
      (NAME: 'OBJTabs';                  PTYPE: StringType;    SVALUE: '2,8,16,18,32,56,80'),
      (NAME: 'PUBPRITabs';               PTYPE: StringType;    SVALUE: '2,4,6,8,10,32,56,80'),
      (NAME: 'DATTabs';                  PTYPE: StringType;    SVALUE: '8,14,24,32,48,56,80'),
      (NAME: 'ExplorerVisible';          PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'EditorPos';                PTYPE: StringType;    SVALUE: '0000,0000,0800,0600,0000'),
      (NAME: 'ExplorerWidth';            PTYPE: IntegerType;   IVALUE: 250),
      (NAME: 'ExplorerPanelSplitPos';    PTYPE: IntegerType;   IVALUE: 150),
      (NAME: 'FileSplitPos';             PTYPE: IntegerType;   IVALUE: 250),
      (NAME: 'FindReplacePos';           PTYPE: StringType;    SVALUE: '0000,0000,0000'),
      (NAME: 'CharChartPos';             PTYPE: StringType;    SVALUE: '0000,0000,0000'),
      (NAME: 'PrefsPos';                 PTYPE: StringType;    SVALUE: '0000,0000,0000'),
      (NAME: 'PortListPos';              PTYPE: StringType;    SVALUE: '0000,0000,0000'),
      (NAME: 'LibraryPaths';             PTYPE: StringType;    SVALUE: ''),                {Default is set at run-time}
      (NAME: 'AutoRecover';              PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'UndoAfterSave';            PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'File01';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File02';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File03';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File04';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File05';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File06';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File07';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File08';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File09';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'File10';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'CheckAssoc';               PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ExtSpin';                  PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ExtSpin2';                 PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ExtBinary';                PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ExtEEPROM';                PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ExtFlash';                 PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'Find01';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find02';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find03';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find04';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find05';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find06';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find07';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find08';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find09';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Find10';                   PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace01';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace02';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace03';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace04';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace05';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace06';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace07';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace08';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace09';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'Replace10';                PTYPE: StringType;    SVALUE: ''),
      (NAME: 'TopFile';                  PTYPE: StringType;    SVALUE: ''),
      (NAME: 'ShowBookmarks';            PTYPE: BooleanType;   BVALUE: False),
      (NAME: 'ShowBlockIndentions';      PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ShowLineNumbers';          PTYPE: BooleanType;   BVALUE: False)(*,
      (NAME: 'AutoIndent';               PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'UseTabCharacter';          PTYPE: BooleanType;   BVALUE: False),
      (NAME: 'BackspaceAutoUnindents';   PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'TabType';                  PTYPE: IntegerType;   IVALUE: 3),
      (NAME: 'CreateBackups';            PTYPE: BooleanType;   BVALUE: False),
      (NAME: 'HighlightSyntax';          PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'FavoriteDirectory01';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory02';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory03';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory04';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory05';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory06';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory07';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory08';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory09';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'FavoriteDirectory10';      PTYPE: StringType;    SVALUE: ''),
      (NAME: 'StartupDirectory';         PTYPE: StringType;    SVALUE: 'Last Used'),  {If changed, make appropriate changes to Pref's UpdateStartupDirectoryList and FavoriteName unit as well}
   *),(NAME: 'LastUsedDirectory';        PTYPE: StringType;    SVALUE: ''),
      (NAME: 'ShowRecentOnly';           PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'FilterIdx';                PTYPE: IntegerType;   IVALUE: 0),
      (NAME: 'SingleInstanceOnly';       PTYPE: BooleanType;   BVALUE: True),
      (NAME: 'ResetSignal';              PTYPE: IntegerType;   IVALUE: 0),
      (NAME: 'ResetDelay';               PTYPE: IntegerType;   IVALUE: 90),
      (NAME: 'ResetDelayP2';             PTYPE: IntegerType;   IVALUE: 10),
      (NAME: 'SerialSearch';             PTYPE: StringType;    SVALUE: 'AUTO'),
      (NAME: 'SerialSearchRules';        PTYPE: StringType;    SVALUE: '+*,-*Bluetooth*,-* BT *,-PropScope'),
      (NAME: 'Initialized';              PTYPE: StringType;    SVALUE: ''),
      (NAME: 'SynScheme';                PTYPE: IntegerType;   IVALUE: CustomSchemeID-1),
      (NAME: 'SynSchemePath';            PTYPE: StringType;    SVALUE: ''),
      (NAME: 'SynSchemeModified';        PTYPE: BooleanType;   BVALUE: False),
      (NAME: 'SynRegular';               PTYPE: StringType;    SVALUE: '$000000000,$000FFFFFF,$000FFFFFF,$00'),   //See GlobalSecondary unit for explanation of these fields
      (NAME: 'SynBlocksCONstant';        PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynBlocksVARiable';        PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynBlocksOBJect';          PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynBlocksPUBlic';          PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynBlocksPRIvate';         PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynBlocksDATa';            PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSpinDebug';             PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSpinIADebug';           PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSpinCommands';          PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSpinConditionals';      PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSpinVariables';         PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynAssemblyDebug';         PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynAssemblyInstructions';  PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynAssemblyConditionals';  PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynAssemblyEffects';       PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynAssemblyLiterals';      PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynDirectives';            PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynCommentsCode';          PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynCommentsDocument';      PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynConstantsNumbers';      PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynConstantsPredefined';   PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynStrings';               PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynOperators';             PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynRegisters';             PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynSizes';                 PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynDataAlignments';        PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'SynDataSizes';             PTYPE: StringType;    SVALUE: '$200000000,$200FFFFFF,$200FFFFFF,$10'),
      (NAME: 'NewP1FileTemplate';        PTYPE: StringType;    SVALUE: ''),                                    {Default is set at run-time}
      (NAME: 'NewP2FileTemplate';        PTYPE: StringType;    SVALUE: '')                                     {Default is set at run-time}
    );

implementation

uses
  Global, Compiler;

procedure OpenRegistry(TheKey: HKEY);
{Open the registry and set root key to TheKey}
begin
  Prefs := TRegistry.Create;
  Prefs.RootKey := TheKey;
end;

{------------------------------------------------------------------------------}

procedure CloseRegistry();
{Safely close the registry}
begin
  Prefs.free
end;

{------------------------------------------------------------------------------}

function UpgradePrefsFromRegistry(): Boolean;
{Look for previous Editor's Version's preference settings in the windows registry and load all relevant settings.
 Returns true if sucessful, false otherwise.
 Versions in the registry are changed only if the next revision causes a legacy issue with previous versions.
 IE: Version 1.0 registry settings support both version 1.0 and 1.1 of the software.

 {All contained under HKEY_CURRENT_USER\<RegPath>}

var
  Sucessful: Boolean;
  Idx: Integer;

  procedure ReadSetting;
  begin {Read the preference setting based on current Idx}
    case PrefType(CPrefs[PrefEntity(Idx)].PType) of
      StringType:  if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then  {Value exists, read value from registry}
                   CPrefs[PrefEntity(Idx)].SValue := Prefs.ReadString(CPrefs[PrefEntity(Idx)].Name);
      IntegerType: if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then  {Value exists, read value from registry}
                   CPrefs[PrefEntity(Idx)].IValue := Prefs.ReadInteger(CPrefs[PrefEntity(Idx)].Name);
      BooleanType: if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then  {Value exists, read value from registry}
                   CPrefs[PrefEntity(Idx)].BValue := Prefs.ReadBool(CPrefs[PrefEntity(Idx)].Name);
    end;
  end;

begin
  {Note, if a previous registry version exits, and a preference does not need to be upgraded, we just load it from the
   registry.  Note that upgraded pref items are excluded from reading from the registry (in this routine) and are
   already properly configured in CPrefs by default, so we'll just skip over those.}
  Sucessful := False;
  OpenRegistry(HKEY_CURRENT_USER);  {Open the registry}
  {Open the previous version Propeller Editor Key}
  if Prefs.OpenKey(RegPath+'2.1.0',False) then
    begin {Found 2.1.0, we'll upgrade from here}
    Sucessful := True;
    for Idx := ord(Low(CPrefs)) to ord(High(CPrefs)) do {Read all preference settings from registry except those listed below; they will be set to defaults or upgraded}
      if not (Idx in [ord(CheckAssoc), ord(NewP1FileTemplate), ord(NewP2FileTemplate), ord(LibraryPaths), ord(LastUsedDirectory), ord(ShowRecentOnly)]) then ReadSetting;
    {No conversion on registry exceptions; just use default}
    Prefs.CloseKey;  {Close key and close registry}
    end
  else
    begin {Else, try to open the next previous version Propeller Editor Key}
    if Prefs.OpenKey(RegPath+'1.3.2',False) then
      begin {Found 1.3.2, we'll upgrade from here}
      Sucessful := True;
      for Idx := ord(Low(CPrefs)) to ord(High(CPrefs)) do {Read all preference settings from registry except those listed below; they will be set to defaults or upgraded}
        if not (Idx in [ord(CheckAssoc), ord(NewP1FileTemplate), ord(NewP2FileTemplate), ord(LibraryPaths), ord(LastUsedDirectory), ord(ShowRecentOnly)]) then ReadSetting;
      {No conversion on registry exceptions; just use default}
//      {Convert NewFileTemplate -> NewP1FileTemplate; NewP2FileTemplate will get default setting}
//      if Prefs.ValueExists('NewFileTemplate') then  {Value exists, read value from registry}
//        CPrefs[NewP1FileTemplate].SValue := Prefs.ReadString('NewFileTemplate');
      Prefs.CloseKey;  {Close key and close registry}
      end;
    end;
  CloseRegistry;
  PerformFinalPrefLoading;    
  PrefsHaveChanged := True;   {Set prefs changed flag so we save upon exit}
  UpgradePrefsFromRegistry := Sucessful;
end;

{------------------------------------------------------------------------------}

function LoadPrefsFromRegistry(): Boolean;
{Loads the preference settings from the windows registry}
{Returns true if sucessful, false otherwise}
{All contained under HKEY_CURRENT_USER\<RegPath>}
var
  Idx           : Integer;

    {----------------}

    procedure PerformMinorUpgrade;
    {Perform minor upgrade on current preferences if necessary.
      Details:
        V1.2.5 and prior used SynPlain instead of SynRegular; if SynRegular does not exist, it is copied from SynPlain.
        V1.2.6 and prior did not feature an exclude-PropScope rule for serial port search.  This will take current rules and append '-PropScope' to them.}
    begin
      {Convert SynPlain to SynRegular}
      if not Prefs.ValueExists('SynRegular') then {SynRegular value does not exist, copy it from SynPlain.}
        if Prefs.ValueExists('SynPlain') then CPrefs[SynRegular].SValue := Prefs.ReadString('SynPlain');
      {Update serial search rules, if necessary, to exclude PropScope by default; this will be parsed later by TPortMetrics.ParseSearchRuleString}
      if not VersionInitialized and (pos('PropScope', CPrefs[SerialSearchRules].SValue) = 0) then
        CPrefs[SerialSearchRules].SValue := CPrefs[SerialSearchRules].SValue + ',-PropScope';
    end;

    {----------------}

begin
  Result := False;
  OpenRegistry(HKEY_CURRENT_USER);  {Open the registry}
  {Open the Propeller Editor Key}
  if Prefs.OpenKey(RegPath+EditorRegistryVersion,False) then
    begin
    {Read all preference settings from registry}
    Result := True;
    for Idx := ord(Low(CPrefs)) to ord(High(CPrefs)) do
      begin
        case PrefType(CPrefs[PrefEntity(Idx)].PType) of
          StringType:  if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then {Value exists, read value from registry}
                         CPrefs[PrefEntity(Idx)].SValue := Prefs.ReadString(CPrefs[PrefEntity(Idx)].Name);
          IntegerType: if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then {Value exists, read value from registry}
                         CPrefs[PrefEntity(Idx)].IValue := Prefs.ReadInteger(CPrefs[PrefEntity(Idx)].Name);
          BooleanType: if Prefs.ValueExists(CPrefs[PrefEntity(Idx)].Name) then {Value exists, read value from registry}
                         CPrefs[PrefEntity(Idx)].BValue := Prefs.ReadBool(CPrefs[PrefEntity(Idx)].Name);
        end;
      end;
    {Perform minor upgrade if necessary}
    PerformMinorUpgrade;
    Prefs.CloseKey;  {Close key and close registry}
    end;
  CloseRegistry;
  PerformFinalPrefLoading;
end;

{------------------------------------------------------------------------------}

procedure PerformFinalPrefLoading;
var
  Idx           : Integer;
  TheFile       : String;
begin
  {Build File History String List (includes bookmark history)}
  for Idx := 0 TO 9 do
    begin
    TheFile := CPrefs[PrefEntity(ord(File01)+Idx)].SValue;
    if TheFile <> '' then FileHistory.Add(TheFile);
    end;
  {Update CPrefs syntax items according to SynScheme, if necessary}
  CPrefs[SynScheme].IValue := min(CustomSchemeID, max(0, CPrefs[SynScheme].IValue));         {Limit syntax scheme value}
  if CPrefs[SynScheme].IValue < CustomSchemeID then
    begin {Default scheme selected}
    BackupCustomSynScheme;
    CopySynScheme(CPrefs[SynScheme].IValue);
    end;
end;

{------------------------------------------------------------------------------}

procedure SavePrefsToRegistry();
{Creates or writes to the preference settings in the windows registry for this application.  All contained under HKEY_CURRENT_USER\<RegPath>}
var
  Idx: Integer;
begin
  for Idx := 0 to FileHistory.Count-1 do CPrefs[PrefEntity(ord(File01)+Idx)].SValue := FileHistory[Idx]; {Update File History Preferences}
  for Idx := FileHistory.Count to 9 do CPrefs[PrefEntity(ord(File01)+Idx)].SValue := '';
  OpenRegistry(HKEY_CURRENT_USER);                    {Open the registry}
  Prefs.OpenKey(RegPath+EditorRegistryVersion, True); {Create version key if it doesn't exist}
  {Save all preference settings to the registry within the version key}
  for Idx := ord(Low(CPrefs)) to ord(High(CPrefs)) do
    begin
      case CPrefs[PrefEntity(Idx)].PType of
        StringType: Prefs.WriteString(CPrefs[PrefEntity(Idx)].Name,CPrefs[PrefEntity(Idx)].SValue);
        IntegerType: Prefs.WriteInteger(CPrefs[PrefEntity(Idx)].Name,CPrefs[PrefEntity(Idx)].IValue);
        BooleanType: Prefs.WriteBool(CPrefs[PrefEntity(Idx)].Name,CPrefs[PrefEntity(Idx)].BValue);
      end;
    end;
  Prefs.CloseKey;  {Close this application's Registry Version key and close registry}
  CloseRegistry;
  PrefsHaveChanged := False;
end;

{------------------------------------------------------------------------------}

function VersionInitialized: Boolean;
{Returns True if this version of software has been run (initialized) prior to this session, False otherwise.  If False, it flags this version
as initializing now (VerInitializing = True) and in the registry (Initialized preference).}
var
  Idx : Integer;
  Ver : String;
  Str : String;
begin
  Result := VerInitializing;
  if Result then exit;
  Ver := uppercase(GetVersionInfo(Application.ExeName, viVersion));
  Str := uppercase(CPrefs[Initialized].SValue);
  repeat
    Idx := pos(Ver, Str);
    Result := (Idx > 0) and ((Idx = 1) or (Str[Idx-1] = '/')) and ((Idx + length(Ver) > length(Str)) or (Str[Idx + length(Ver)] = '/'));
    if not Result and (Idx > 0) then delete(Str, 1, Idx + length(Ver) - 1);
  until Result or (Idx = 0) or (Str = '');
  if not Result then
    begin  {Not run (initialized) previous to this session; mark it as initialized now}
    CPrefs[Initialized].SValue := CPrefs[Initialized].SValue + ifthen(length(CPrefs[Initialized].SValue) > 0, '/', '') + GetVersionInfo(Application.ExeName, viVersion);
    PrefsHaveChanged := True;
    VerInitializing := True;
    end;
end;

{------------------------------------------------------------------------------}

procedure CheckFileAssociationsInRegistry(ForceAssociations: Boolean);
{Checks file association settings in the windows registry for the Propeller IDE.  Extensions are contained under
HKEY_CURRENT_USER\Software\Classes\.extension and ProgID is contained in HKEY_CURRENT_USER\Software\Classes\Propeller.SourceCode.1.}
var
  Idx        : PrefEntity;
  Changed    : Boolean;
  Response   : Integer;
  RegWritten : Boolean;
  AssocRoot  : String;
  Temp       : String;
const
  Create = True;
  ProgID = 'Propeller.SourceCode.1';
begin
  Response := 0;
  RegWritten := False;
  {Exit if not asked to check or force associations, and this version of software has previously run (initialized)}
  {$IFNDEF SX_TESTER_AS_PROGRAMMER} if VersionInitialized and not (CPrefs[CheckAssoc].BValue or ForceAssociations) then {$ENDIF} exit;  {Exit if not asked to check or set associations; or if SX_TESTER_AS_PROGRAMMER build}
  {Open the registry in Windows 2K/XP}
  OpenRegistry(HKEY_CURRENT_USER);
  Prefs.OpenKey('\Software\Classes',False);
  AssocRoot := '\Software\Classes\';
  try
    Changed := ForceAssociations;
    {Check existing associations, if necessary}
    if not ForceAssociations then
      begin  {If we're not supposed to write associations regardless, then check them first to see if they're already set}
      for Idx := ExtSpin to ExtFlash do
        begin {Check extensions}
        if CPrefs[Idx].BValue then
          begin {Extension is allowed to be associated}
          case Idx of
            ExtSpin   : Temp := AssocRoot+'.'+Prop1SourceExt;
            ExtSpin2  : Temp := AssocRoot+'.'+Prop2SourceExt;
            ExtBinary : Temp := AssocRoot+'.'+Prop1AppBinExt;                        {P2 binary extension matches P1}
            ExtEEPROM : Temp := AssocRoot+'.'+Prop1AppEEExt;
            ExtFlash  : Temp := AssocRoot+'.'+Prop2AppFLExt;
          end;
          if Prefs.KeyExists(Temp) then
            begin
            Prefs.OpenKey(Temp, False);
            if not (Prefs.ReadString('') = ProgID) and not Changed then Changed := True {Check .ext key's (Default) value}
            end
          else
            Changed := True;
          end;
        end;  {Check extensions}
      if Prefs.KeyExists(AssocRoot+ProgID+'\shell\open\command') then
        begin {Check version of application in progID}
        Prefs.OpenKey(AssocRoot+ProgID+'\shell\open\command', false);
        Temp := Prefs.ReadString('');
        if Temp <> '' then
          begin
          if Temp[1] = '"' then {Get exe name, if it is quoted}
            Temp := copy(Temp, 2, pos('"', copy(Temp, 2, length(Temp)-1))-1)
          else                  {Get exe name, if it is not quoted}
            Temp := copy(Temp, 1, pos(' ', copy(Temp, 1, length(Temp)-1)));
          end;
        if (GetVersionInfo(Temp, viVersion) <> GetVersionInfo(Application.ExeName, viVersion)) and not Changed then Changed := True;
        end
      else
        Changed := True;
      if Prefs.KeyExists(AssocRoot+ProgID+'\DefaultIcon') then
        begin {Check dafault icon for source file}
        Prefs.OpenKey(AssocRoot+ProgID+'\DefaultIcon', false);
        Temp := Prefs.ReadString('');
        if Temp <> '' then
          begin
          if Temp[1] = '"' then {Get exe name, if it is quoted}
            Temp := copy(Temp, 2, pos('"', copy(Temp, 2, length(Temp)-1))-1)
          else                  {Get exe name, if it is not quoted}
            Temp := copy(Temp, 1, pos(' ', copy(Temp, 1, length(Temp)-1)));
          end;
        if (GetVersionInfo(Temp, viVersion) <> GetVersionInfo(Application.ExeName, viVersion)) and not Changed then Changed := True;
        end
      else
        Changed := True;
      end;
    {If associations have been altered, were not set at all or we've been asked to write them anyway...}
    if Changed or ForceAssociations then
      begin {If registry file associations not correct, and we've been asked to check, notify user}
      if not ForceAssociations then
        begin
        Temp := #$D#$A;
        for Idx := ExtSpin to ExtFlash do
          if CPrefs[Idx].BValue then {Extension is allowed to be associated}
            case Idx of
              ExtSpin   : Temp := Temp + #$D#$A#$9 + '.'+Prop1SourceExt;
              ExtSpin2  : Temp := Temp + #$D#$A#$9 + '.'+Prop2SourceExt;             
              ExtBinary : Temp := Temp + #$D#$A#$9 + '.'+Prop1AppBinExt;             {P2 binary extension matches P1}
              ExtEEPROM : Temp := Temp + #$D#$A#$9 + '.'+Prop1AppEEExt;
              ExtFlash  : Temp := Temp + #$D#$A#$9 + '.'+Prop2AppFLExt;
            end;
        Temp := Temp + #$D#$A#$D#$A;
        messagebeep(MB_ICONINFORMATION);
        Response := MessageDlg('One or more of the following Propeller file types are not associated with '+
                               'this '+PropIDEName+'.'+Temp+'Associate file type(s) now?'+#13+#10+#13+#10+
                               'Note: select Cancel to stop checking file associations at startup.', mtConfirmation,
                               [mbYes, mbNo, mbCancel], 0);
        end;
      if (Response = mrYes) or ForceAssociations then
        begin {User selected Yes or we've been asked to write them anyway}
        for Idx := ExtSpin to ExtFlash do
          begin {Associate extensions}
          if CPrefs[Idx].BValue then
            begin {Extension is allowed to be associated}
            case Idx of
              ExtSpin   : Temp := AssocRoot+'.'+Prop1SourceExt;
              ExtSpin2  : Temp := AssocRoot+'.'+Prop2SourceExt;
              ExtBinary : Temp := AssocRoot+'.'+Prop1AppBinExt;                      {P2 binary extension matches P1}
              ExtEEPROM : Temp := AssocRoot+'.'+Prop1AppEEExt;
              ExtFlash  : Temp := AssocRoot+'.'+Prop2AppFLExt;
            end;
            Prefs.OpenKey(Temp, Create);                                             {Create the .ext key(s)}
            Prefs.WriteString('', ProgID);
            end;
          end;
        {Set ProgID}
        Prefs.OpenKey(AssocRoot+ProgID, True);                                       {Create the Propeller.SourceCode.1 key}
        Prefs.WriteString('','Propeller Source Code');
        Prefs.OpenKey('shell\open',True);                                            {Create the shell\open key}
        Prefs.WriteString('','Open with &'+PropIDEName);
        Prefs.OpenKey('command',True);                                               {Create the command key}
        Prefs.WriteString('','"'+Application.ExeName+'" "%1"');                      {Write command-line string}
        Prefs.OpenKey(AssocRoot+ProgID, False);                                      {Go up to the Propeller.SourceCode.1 key}
        Prefs.OpenKey('DefaultIcon',True);                                           {Create the DefaultIcon key}
        Prefs.WriteString('','"'+Application.ExeName+'",'+inttostr(DefaultIconIdx)); {Write command-line string}
        Prefs.CloseKey;  {Close Propeller IDE key and close registry}
        RegWritten := True;
        end  {User selected Yes}
      else
        if Response = mrCancel then
          begin  {User selected Cancel, turn off association checks upon startup}
          PrefsHaveChanged := True;
          CPrefs[CheckAssoc].BValue := False;
          end;
      end;
  finally
    CloseRegistry;
    if RegWritten then SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nil, nil); {Notify system of file association change}
  end;
end;

{------------------------------------------------------------------------------}
(*
procedure ReadPrefs(Parent: TWinControl; Prefs: array of TPrefs);
{This procedure iterates through all controls and child controls in Parent and
updates those that contain preferences with the corresponding current preference
setting.  Only controls who's .Tag is 10, and who appear in the PrefEntity set are
modified.  This procedure works recursively to traverse the family tree starting
with the Parent.  As written, only TShape, TComboBox, TEdit, TRadioButton and
TCheckBox are supported.  Unfortunately, it had to be written this way because the
Text, Brush and Checked properties do not exist or were not published in the
ancestor classes}
var
  CIdx: Integer;
  PIdx: Integer;
begin
  for CIdx := 0 to Parent.ControlCount-1 do
  begin
    {If current control is of type TWinControl and it has children, recursively iterate throgh the children}
    if (Parent.Controls[CIdx] is TWinControl) and (TWinControl(Parent.Controls[CIdx]).ControlCount > 0) then
      ReadPrefs(TWinControl(Parent.Controls[CIdx]), Prefs)
    else
      begin
      if Parent.Controls[CIdx].Tag = 10 then
        {A preference control was found, look for it's matching preference setting and set it}
        begin
          PIdx := Low(Prefs);
          while (PIdx <= High(Prefs)) and (Prefs[PIdx].Name <> Parent.Controls[CIdx].Name) do
            inc(PIdx);
          case Prefs[PIdx].PType of
            {Current setting is a StringType value}
            StringType:  begin
                           if Parent.Controls[CIdx] is TComboBox then
                             TComboBox(Parent.Controls[CIdx]).ItemIndex :=
                               TComboBox(Parent.Controls[CIdx]).Items.IndexOf(Prefs[PIdx].SValue)
                           else
                             if Parent.Controls[CIdx] is TEdit then
                               TEdit(Parent.Controls[CIdx]).Text := Prefs[PIdx].SValue;
                         end;
            {Current setting is an IntegerType value}
            IntegerType: begin
                         if Parent.Controls[CIdx] is TShape then
                           TShape(Parent.Controls[CIdx]).Brush.Color := Prefs[PIdx].IValue
                         else
                           if Parent.Controls[CIdx] is TComboBox then
                             TComboBox(Parent.Controls[CIdx]).ItemIndex := (Prefs[PIdx].IValue);
                         end;
           {Current setting is a BooleanType value}
            BooleanType: begin
                         if Parent.Controls[CIdx] is TRadioButton then
                           TRadioButton(Parent.Controls[CIdx]).Checked := Prefs[PIdx].BValue
                         else
                           if Parent.Controls[CIdx] is TCheckBox then
                             TCheckBox(Parent.Controls[CIdx]).Checked := Prefs[PIdx].BValue;
                         end;
          end;
        end;
      end;
  end;
end;

{------------------------------------------------------------------------------}

procedure WritePrefs(Parent: TWinControl);
{This procedure iterates through all controls and child controls in Parent and
updates the CPrefs preference settings to what's contained in the controls.  These
values are NOT written to the registry by this routine, WritePrefsToRegistry does
that.  Only controls who's .Tag is 10, and who appear in the PrefEntity set are
viewed.  This procedure works recursively to traverse the family tree starting
with the Parent.  As written, only TShape, TComboBox, TEdit, TRadioButton and
TCheckBox are supported.  Unfortunately, it had to be written this way because the
Text, Brush and Checked properties do not exist or were not published in the
ancestor classes}
var
  CIdx: Integer;
  PIdx: Integer;
begin
  for CIdx := 0 to Parent.ControlCount-1 do
    begin
    {If current control is of type TWinControl and it has children, recursively iterate throgh the children}
    if (Parent.Controls[CIdx] is TWinControl) and (TWinControl(Parent.Controls[CIdx]).ControlCount > 0) then
      WritePrefs(TWinControl(Parent.Controls[CIdx]))
    else
      begin
      if Parent.Controls[CIdx].Tag = 10 then
        {A preference control was found, look for it's matching preference setting and set it}
        begin
          PIdx := ord(Low(CPrefs));
          while (PIdx <= ord(High(CPrefs))) and (CPrefs[PrefEntity(PIdx)].Name <> Parent.Controls[CIdx].Name) do
            inc(PIdx);
          case CPrefs[PrefEntity(PIdx)].PType of
           {Current setting is a StringType value}
            StringType:  begin
                           if Parent.Controls[CIdx] is TComboBox then
                             CPrefs[PrefEntity(PIdx)].SValue := TComboBox(Parent.Controls[CIdx]).Text
                           else
                             if Parent.Controls[CIdx] is TEdit then
                               CPrefs[PrefEntity(PIdx)].SValue := TEdit(Parent.Controls[CIdx]).Text;
                         end;
            {Current setting is an IntegerType value}
            IntegerType: begin
                         if Parent.Controls[CIdx] is TShape then
                           CPrefs[PrefEntity(PIdx)].IValue := TShape(Parent.Controls[CIdx]).Brush.Color
                         else
                           if Parent.Controls[CIdx] is TComboBox then
                             CPrefs[PrefEntity(PIdx)].IValue := TComboBox(Parent.Controls[CIdx]).ItemIndex;
                         end;
           {Current setting is a BooleanType value}
            BooleanType: begin
                         if Parent.Controls[CIdx] is TRadioButton then
                           CPrefs[PrefEntity(PIdx)].BValue := TRadioButton(Parent.Controls[CIdx]).Checked
                         else
                           if Parent.Controls[CIdx] is TCheckBox then
                             CPrefs[PrefEntity(PIdx)].BValue := TCheckBox(Parent.Controls[CIdx]).Checked;
                         end;
          end;
        end;
      end;
    end;
end;
*)

{------------------------------------------------------------------------------}

function GetPosValue(PosStr: String; Position: Integer): Integer;
{Extract the position value from a window position string.
Pos: 0 = Left, 1 = Top, 2 = Right/Used, 3 = Bottom, 4 = Maximized, 5 = SplitterBar}
var
  Idx   : Integer;
  Count : Integer;
begin
  Idx := 1;
  for Count := 0 to Position-1 do Idx := Idx + pos(',',copy(PosStr,Idx,length(PosStr)-Idx+1));
  Result := StrToInt(copy(PosStr,Idx,6));
end;

{------------------------------------------------------------------------------}

function PutPosValue(Pos: array of Integer): String;
{Insert the position values from a window's position into a position string.}
var
  Idx: Integer;
  ResultStr: String;
begin
  ResultStr := '';
  for Idx := 0 to high(Pos) do ResultStr := ResultStr + format('%.5d,',[Pos[Idx]]);
  Result := copy(ResultStr,1,length(ResultStr)-1);
end;

{------------------------------------------------------------------------------}

procedure LoadWindowMetrics(Window: TForm; Pref: PrefEntity);
{Set window to metrics recorded in preferences.  Automatically move position if loaded position would make its title bar more than 50% outside visible desktop.}
var
  PrefRect : TRect;
  PrefMax  : Integer;
  Metrics  : Integer;
begin
  {Get preference values}
  PrefRect.Left   := GetPosValue(CPrefs[Pref].SValue, 0);
  PrefRect.Top    := GetPosValue(CPrefs[Pref].SValue, 1);
  PrefRect.Right  := GetPosValue(CPrefs[Pref].SValue, 2);
  PrefRect.Bottom := GetPosValue(CPrefs[Pref].SValue, 3);
  PrefMax         := GetPosValue(CPrefs[Pref].SValue, 4);
  {Determine number of metrics for Window}
  Metrics := ifthen(Window <> Application.MainForm, 2, 4);
  {Exit if no metrics recorded (ie: first time run / preference set to default}
  if (Metrics = 2) and (PrefRect.Right = 0) then exit;
  {Adjust .Right and .Bottom if those metrics were not saved (Metrics = 2)}
  if Metrics = 2 then PrefRect.BottomRight := point(PrefRect.Left+Window.Width, PrefRect.Top+Window.Height);
  {Ensure window will be visible}
  PrefRect := GetDisplayableWindowPosition(PrefRect);
  {Set to last known position/size}
  WinPos.flags := 0;
  WinPos.ptMaxPosition := Point(0,0);
  WinPos.rcNormalPosition := PrefRect;
  WinPos.showCmd := ifthen(Metrics = 2, SW_HIDE, ifthen(PrefMax = ord(wsMaximized), SW_SHOWMAXIMIZED, SW_SHOWNORMAL));
  SetWindowPlacement(Window.Handle, WinPos);
end;

{------------------------------------------------------------------------------}

procedure SaveWindowMetrics(Window: TForm; Pref: PrefEntity);
{Retrieve and save window metrics to preferences}
var
  PLeft, PTop, PRight, PBottom, PMax : Integer;
  Metrics                            : Integer;
begin
  {Get preference values}
  PLeft   := GetPosValue(CPrefs[Pref].SValue, 0);
  PTop    := GetPosValue(CPrefs[Pref].SValue, 1);
  PRight  := GetPosValue(CPrefs[Pref].SValue, 2);
  PBottom := GetPosValue(CPrefs[Pref].SValue, 3);
  PMax    := GetPosValue(CPrefs[Pref].SValue, 4);
  {Determine number of metrics for Window}
  Metrics := ifthen(Window <> Application.MainForm, 2, 4);
  {Save metrics if necessary}
  GetWindowPlacement(Window.Handle, WinPos);
  with WinPos.rcNormalPosition do
    begin
    if ((Metrics >= 0) and (PLeft   <> Left  )) or                 {Left}
       ((Metrics >= 1) and (PTop    <> Top   )) or                 {Top}
       ((Metrics =  2) and (PRight  <> 1     )) or                 {Used (flag)}
       ((Metrics >  2) and (PRight  <> Right )) or                 {Right}
       ((Metrics >= 3) and (PBottom <> Bottom)) or                 {Bottom}
       ((Metrics >= 4) and (ord(Window.WindowState) <> PMax)) then {Maximimzed/Normal}
      begin
      PrefsHaveChanged := True;
      case Metrics of
        2 : CPrefs[Pref].SValue := PutPosValue([Left, Top, 1]); {Note: width marked with "used" flag}
        4 : CPrefs[Pref].SValue := PutPosValue([Left, Top, Right, Bottom, ord(Window.WindowState)]);
      end;
      end;
    end;
end;

{------------------------------------------------------------------------------}

function GetDisplayableWindowPosition(Bounds: TRect): TRect;
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

procedure EnsureWindowDisplayable(Window: TForm);
{Ensure window is reasonably within displayable coordinates.  Moves window if necessary.}
begin
  GetWindowPlacement(Window.Handle, WinPos);
  WinPos.rcNormalPosition := GetDisplayableWindowPosition(Window.BoundsRect);
  if not (PointsEqual(WinPos.rcNormalPosition.TopLeft, Window.BoundsRect.TopLeft) and PointsEqual(WinPos.rcNormalPosition.BottomRight, Window.BoundsRect.BottomRight)) then
    SetWindowPlacement(Window.Handle, WinPos);
end;

{------------------------------------------------------------------------------}

Initialization
  {Initialize CPrefs array with default values in case this is the first time this program is run}
  for PIdx := Low(CPrefs) to High(CPrefs) do CPrefs[PIdx] := DPrefs[PIdx];
  PrefsHaveChanged := False;
  VerInitializing := False;
  {Allocate memory for Window Placement structure and set its length}
  GetMem(WinPos,sizeof(TWindowPlacement));
  WinPos.length := sizeof(TWindowPlacement);

Finalization
  {Free WinPos memory}
  freemem(WinPos);

end.
