!Problem: Pasted URLs can have Characters > 80h that Zip rejects. 
!This happens dragging and dropping URL from IE
!These are usually .URL or .WEBSITE files
!------------------------------------------------------------------------------
! 25-Mar-2021  Added Shell Auto Complete of File Name SHAutoComplete() 
!              from https://github.com/CarlTBarnes/AutoCompleteFileOrURL
!------------------------------------------------------------------------------
  PROGRAM
  INCLUDE('KEYCODES.CLW')
  MAP
ZipFileNameFixer    PROCEDURE()
DB                  PROCEDURE(STRING xMessage)
   MODULE('API')
     OutputDebugString(*CSTRING),RAW,PASCAL,NAME('OutputDebugStringA'),DLL(1)
     SHAutoComplete(SIGNED hwndEdit, LONG dwFlags ),PASCAL,DLL(1),LONG,PROC  !Returns HResult < 0 if failed
     !MSDN: "WARNING: Caller needs to have called CoInitialize() or OleInitialize()".
     !Must call CoInit if calling SHAutoComplete() before Event:OpenWindow (probably RTL calls by then). Must be called once per thread
     SHAuto_CoInitialize(long=0),PASCAL,DLL(1),LONG,PROC,NAME('CoInitialize')     
   END  
  END

SHACF_FILESYS_DIRS  EQUATE(00000020h)  !Only directories, UNC servers, and UNC server shares. No files.
MsgCaption  STRING('Zip File Name Fixer - Scan for 80h-FFh')
CfgFileIni  STRING('.\ZipFNFix.INI')
CfgSection  EQUATE('Config')

LogFile     FILE,DRIVER('BASIC','/FirstRowHeader=ON'),NAME('ZipFNFixLog.csv'),PRE(Log),CREATE
Record        RECORD,PRE()
RenDate         STRING(@D10-)
RenTime         STRING(@t04)  
BadChar         STRING(13) 
Path            STRING(255)
OldName         STRING(255)
NewName         STRING(255)
              END 
            END
          
  CODE
  ZipFileNameFixer()
  
ZipFileNameFixer  PROCEDURE()
pFolder PSTRING(256)

QNdx    LONG,AUTO
DirQ    QUEUE(FILE:Queue),PRE(DirQ)
NewName         STRING(255)
PathBS          STRING(255)     !always has \
        END ! DirQ:Name  DirQ:ShortName(8.3?)  DirQ:Date  DirQ:Time  DirQ:Size  DirQ:Attrib  DirQ:PathBS

!--Data declarations needed to recurse directories
Dirs2LoadQ      QUEUE              !Directories to load, added to as they are found
Name                STRING(255)
                END
Dirs2Ndx        LONG,AUTO          !Current Record in Dirs2LoadQ to load for GET(,Dirs2Ndx)
Dirs2AddSpot    LONG,AUTO          !For ADD(Dirs2Load2,Dirs2AddSpot) so dirs are inserted in a good order
Dir1Loading     STRING(255),AUTO   !Current directory loading (from Dirs2LoadQ)
Dir1Q           QUEUE(FILE:Queue),PRE(Dir1Q).   !The current Dir (Dir1Loading) files

FolderInput     STRING(255)
ScanAllExt      BYTE
ScanSubs        BYTE(1)

ErrorExample STRING('This utility fixes the below Compress Error by renaming files ' &|
     'that contain ASCII 128-255 or "?" Unicode characters in the name. ' &|
     'These are normally a result of creating web page shortcuts that use the page title for the name.' &| 
     '<13,10>={60}' &|
     '<13,10>Compressed (zipped) Folders Error' &|
     '<13,10>' &|
     '<13,10>''D:\Path\Page ? blarg <135> urr.url'' cannot be compressed because it ' &|
     'includes characters that cannot be used in a compressed folder, such ' &|
     'as ?. You should rename this file or directory.')
     
Dirs2Window WINDOW('Scan for Zip Problems 80h-FFh'),AT(,,340,153),CENTER,GRAY,SYSTEM,ICON(ICON:Pick),FONT('Segoe UI',9)
        PROMPT('&Folder:'),AT(2,4),USE(?Prompt:Folder)
        ENTRY(@s255),AT(26,4,289,11),USE(FolderInput),FONT('Consolas') 
        BUTTON('...'),AT(319,3,12,12),USE(?PickDirBtn),SKIP
        CHECK('Subfolders'),AT(217,23),USE(ScanSubs),SKIP,TIP('Scan subfolders')
        CHECK('&All Files (uncheck *.URL and *.WebSite)'),AT(79,23),USE(ScanAllExt),SKIP,TIP('Normally scan only .URL an' & |
                'd .WEBSITE extensions')
        BUTTON('&Scan'),AT(25,19,45),USE(?ScanBtn)
        BUTTON('&Close'),AT(271,19,45),USE(?CloseBtn)
        STRING(@s255),AT(25,41,291),USE(Dir1Loading)
        PROGRESS,AT(25,53,291,8),USE(?Dirs2Progress),RANGE(0,100)
        TEXT,AT(25,70,291,75),USE(ErrorExample),SKIP,FLAT,VSCROLL,READONLY
    END

FilesWindow WINDOW('Directory'),AT(,,613,250),GRAY,SYSTEM,ICON(ICON:Pick),FONT('Segoe UI',9,,FONT:regular),RESIZE
        BUTTON('Find Problems'),AT(1,1),USE(?FindProblemsBtn),TIP('Filter to show only problem file names.<13,10>Problem' & |
                ' Characters will be in Short Name.<13,10>New Name column will be filled in')
        BUTTON('Fix by Rename'),AT(63,1),USE(?FixProblemsBtn),DISABLE,TIP('Rename the problem files to New Name')
        BUTTON('Debug1'),AT(142,1),USE(?Debug1Btn),SKIP,TIP('For selected file on clipboard list problem characters')
        STRING('This will find files with Characters 80h-FFh that cause trouble for ZIP and CD Burn, then rename them'), |
                AT(199,1)
        STRING('You can double click on files to open in Windows Explorer.'),AT(199,10)
        LIST,AT(1,21),FULL,USE(?ListFiles),VSCROLL,TIP('Click heads to sort, Right Click to reverse, Ctrl+C to Copy'), |
                FROM(DirQ),FORMAT('155L(1)|M~Name (double click on files for info)~@s255@52L(1)|M~Short Name~@s13@28R(2)' & |
                '|M~Date~C(0)@d1@28R(2)|M~Time~C(0)@T4@45R(1)|M~Size~C(0)@n13@16R(1)|M~Attr~L(1)@n3@127L(1)|M~New Name (' & |
                'fixed 80h-FFh)~@s255@100L(1)|M~Path\~@s255@'),ALRT(CtrlC), ALRT(DeleteKey)
    END
Dir2Clp     ANY
MsgTxt      ANY
WX  LONG
WY  LONG
SortNow  LONG
SortLast LONG 

DOO     CLASS
ConfigLoad    PROCEDURE()
ConfigSave    PROCEDURE()
ProblemsFind  PROCEDURE()
FixProblems   PROCEDURE(),BOOL
LogFileOpen   PROCEDURE(),BOOL
LogFileAdd    PROCEDURE()
LogFileClose  PROCEDURE()
Debug1  PROCEDURE()
        END
    CODE
    SYSTEM{PROP:PropVScroll}=1 ; SYSTEM{PROP:MsgModeDefault}=MSGMODE:CANCOPY
    FREE(DirQ)
    DOO.ConfigLoad()
    pFolder=''
    !--Recurse directories, load (*.URL;*.WEBSITE) files into DirQ
    Dir1Loading = 'Enter your Folder and press Scan'
    OPEN(Dirs2Window)
    ACCEPT
      CASE EVENT()
      OF EVENT:OpenWindow   
            SHAuto_CoInitialize()   !RTL has called but to be safe call
            SHAutoComplete(?FolderInput{PROP:Handle},SHACF_FILESYS_DIRS)
      OF EVENT:CloseWindow ; RETURN
      END
      CASE ACCEPTED()
      OF ?CloseBtn ; RETURN
      OF ?PickDirBtn ; FILEDIALOG('Folder to Scan',FolderInput,,FILE:Directory+FILE:LongName+FILE:KeepDir) ; DISPLAY
      OF ?ScanBtn
            FolderInput=LEFT(FolderInput)
            QNdx=LEN(CLIP(FolderInput)) ; IF FolderInput[QNdx]='\' THEN FolderInput[QNdx]=''.
            IF ~FolderInput OR ~EXISTS(FolderInput) THEN
                Dir1Loading = 'Folder entered does not exist!'
                SELECT(?FolderInput) ; BEEP ; CYCLE
            END
            DOO.ConfigSave()
            pFolder=CLIP(FolderInput)
            FREE(Dirs2LoadQ) ; FREE(Dir1Q)
            Dirs2LoadQ=pFolder & '\'
            ADD(Dirs2LoadQ)
            Dirs2Ndx=0
            DISABLE(?ScanBtn)
            Dir1Loading = 'Loading...'
            0{PROP:Timer}=1
      END
      CASE EVENT()
      OF EVENT:Timer
        Dirs2Ndx += 1
        GET(Dirs2LoadQ,Dirs2Ndx)
        IF ERRORCODE() THEN 
           0{PROP:Timer}='' 
           IF ~Records(DirQ) THEN 
               Message('No files found in ' & FolderInput,MsgCaption,ICON:Asterisk)
               ?Dirs2Progress{PROP:Progress}=0
               ENABLE(?ScanBtn)
               Dir1Loading = 'No files found in ' & FolderInput
               CYCLE
           END                    
           BREAK
        END
        Dir1Loading=Dirs2LoadQ
        Dirs2AddSpot=Dirs2Ndx + 1
        DISPLAY
        DIRECTORY(Dir1Q,CLIP(Dir1Loading)&'*.*',ff_:NORMAL+ff_:DIRECTORY)
        LOOP
            GET(Dir1Q,1)
            IF ERRORCODE() THEN BREAK.
            DELETE(Dir1Q)
            IF BAND(Dir1Q:Attrib,FF_:Directory)
               IF Dir1Q:Name='.' OR Dir1Q:Name='..' THEN CYCLE.
               IF ~ScanSubs THEN CYCLE.
               Dirs2LoadQ=CLIP(Dir1Loading) & CLIP(Dir1Q:Name) & '\'
               ADD(Dirs2LoadQ,Dirs2AddSpot)
               Dirs2AddSpot += 1
               CYCLE !Only put Files into DirQ
            END
            IF  ~ScanAllExt                                            |
            AND ~MATCH(Dir1Q.Name,'*.URL',MATCH:Wild+Match:NoCase)     |
            AND ~MATCH(Dir1Q.Name,'*.WEBSITE',MATCH:Wild+Match:NoCase) | 
                     THEN CYCLE.
    
            DirQ :=: Dir1Q
            DirQ:PathBS = Dir1Loading
            DirQ:Name=UPPER(DirQ:Name[1]) & DirQ:Name[2 : SIZE(DirQ:Name) ]
            DirQ:PathBS=UPPER(DirQ:PathBS[1]) & DirQ:PathBS[2 : SIZE(DirQ:PathBS) ]
            ADD(DirQ)
        END
        ?Dirs2Progress{PROP:Progress}=Dirs2Ndx / RECORDS(Dirs2LoadQ) * 100
      END !CASE EVENT()
    END !LOOP thru Dirs2LoadQ
    FREE(Dirs2LoadQ) !List of paths loaded can be freed if you don't need it
    GETPOSITION(0,WX,WY)
    CLOSE(Dirs2Window)
    IF ~Records(DirQ) THEN 
        Message('Ended with No files found in ' & FolderInput,MsgCaption)
        RETURN
    END
    !--END Recurse directories, DirQ has files loaded
    SORT(DirQ , DirQ:PathBS, DirQ:Name )  !   , DirQ:PathBS , DirQ:ShortName , DirQ:Date , DirQ:Time , DirQ:Size , DirQ:Attrib )
    
    !===== Test Window to View Directory() ========================================
    OPEN(FilesWindow)
    SETPOSITION(0,WX,WY)

    ?ListFiles{PROP:NoTheme}=1             !If Manifested allows header colors
    ?ListFiles{PROPLIST:HasSortColumn}=1                                    !Colors  are OPTIONAL
    ?ListFiles{PROPLIST:HdrSortBackColor} = COLOR:HIGHLIGHT                 !Color Header Background    
    ?ListFiles{PROPList:HdrSortTextColor} = COLOR:HIGHLIGHTtext             !Color Header Text
    ?ListFiles{PROPList:SortBackColor}    = 80000018h !COLOR:InfoBackground !Color List Background
    ?ListFiles{PROPList:SortTextColor}    = 80000017h !COLOR:InfoText       !Color List Text   
    
    0{PROP:Text}='Directory ' & CHOOSE(~ScanAllExt,'*.Url *.Website','*.*') & |
                 ' (' & RECORDS(DirQ) &' records) ' & pFolder
    ACCEPT
      CASE ACCEPTED()
      OF ?FindProblemsBtn ; DOO.ProblemsFind() ; DISPLAY
                            0{PROP:Text}='Zip Problems (' & RECORDS(DirQ) &' records) ' & pFolder
                            IF RECORDS(DirQ) THEN ENABLE(?FixProblemsBtn).
                            DISABLE(?FindProblemsBtn)    
                            
      OF ?FixProblemsBtn  ; IF ~DOO.FixProblems() THEN CYCLE.
                            DISPLAY       
                            0{PROP:Text}='Fixed Problems (' & RECORDS(DirQ) &' records) ' & pFolder
                            DISABLE(?FixProblemsBtn)

      OF ?Debug1Btn ; DOO.Debug1()
      END
    
       CASE FIELD()
       OF ?ListFiles
          CASE EVENT() !Click Header to Sort, Right Click Reverses
          OF EVENT:NewSelection !DblClick to View 1 File
             IF KEYCODE()=MouseLeft2 THEN
                GET(DirQ,CHOICE(?ListFiles))
                MsgTxt=CLIP(DirQ:Name) &'<13,10,13,10>Short: '& DirQ:ShortName & |
                        '<13,10>Date:<9> '& DirQ:Date &'<13,10>Time:<9> '& DirQ:Time &'<13,10>Size:<9> '& DirQ:Size & |
                        '<13,10>Attr:<9> '& DirQ:Attrib &'<13,10>Path: '& CLIP(DirQ:PathBS) & |
                        '<13,10,13,10>New Name: ' & CLIP(DirQ:Name) & |
                        '<13,10,13,10>Length Path+Name: ' & LEN( CLIP(DirQ:PathBS) & CLIP(DirQ:Name) ) &'  (Limit 250)'
                CASE Message(MsgTxt,'File '&CHOICE(?ListFiles) &' - '&MsgCaption,,'Close|Explore|Copy|Debug 1')
                OF 3 ; SETCLIPBOARD(MsgTxt) 
                OF 2 ; SETCLIPBOARD( CLIP(DirQ:PathBS) ) 
                       IF INSTRING('?',DirQ:Name) THEN    !Path has Unicode, then file name has ? andis bad
                          RUN('Explorer.exe /n,"' & CLIP(DirQ:PathBS) &'"') 
                       ELSE    
                          RUN('Explorer.exe /select,"' & CLIP(DirQ:PathBS) & CLIP(DirQ:Name) &'"') 
                       END
                OF 4 ; DOO.Debug1()
                END
             END
          OF EVENT:HeaderPressed
                SortNow=?ListFiles{PROPList:MouseDownField}    !This is really Column and NOT Q Field
                SortNow=?ListFiles{PROPLIST:FieldNo,SortNow}   !Now we have Queue Field to use with WHO()
                SORT(DirQ,CHOOSE(SortNow<>SortLast,'+','-') & WHO(DirQ,SortNow))
                SortLast=CHOOSE(SortNow<>SortLast,SortNow,SortLast * -1)
                DISPLAY
          
          OF EVENT:PreAlertKey
             CASE KEYCODE()
             OF DeleteKey ; GET(DirQ,CHOICE(?ListFiles)) ; DELETE(DirQ) ; DISPLAY
             OF CtrlC
                Dir2Clp='Name<9>Short<9>Date<9>Time<9>Size<9>Attr<9>Path'
                LOOP QNdx = 1 TO RECORDS(DirQ)
                   GET(DirQ,QNdx)
                   Dir2Clp=Dir2Clp & CLIP('<13,10>' & |
                           CLIP(DirQ:Name) &'<9>'& DirQ:ShortName &'<9>'& FORMAT(DirQ:Date,@d02) &'<9>'& FORMAT(DirQ:Time,@t04) |
                           &'<9>'& DirQ:Size &'<9>'& DirQ:Attrib &'<9>'& DirQ:PathBS)
                END 
                SETCLIPBOARD(Dir2Clp)
             END !IF CtrlC
          END ! Case EVENT()
       END ! Case FIELD()
    END !ACCEPT
    CLOSE(FilesWindow)
    RETURN
!-----------------------------
DOO.ConfigLoad    PROCEDURE()
    CODE
    FolderInput=GETINI(CfgSection,'Path',FolderInput,CfgFileIni)
    ScanSubs   =GETINI(CfgSection,'Subs',ScanSubs   ,CfgFileIni)
    ScanAllExt =GETINI(CfgSection,'All' ,ScanAllExt ,CfgFileIni)

DOO.ConfigSave    PROCEDURE()
    CODE
    PUTINI(CfgSection,'Path',FolderInput,CfgFileIni)
    PUTINI(CfgSection,'Subs',ScanSubs   ,CfgFileIni)
    PUTINI(CfgSection,'All' ,ScanAllExt ,CfgFileIni)
                        
!-----------------------------
DOO.ProblemsFind  PROCEDURE()
BadChr  BYTE,DIM(256) 
X       USHORT
LnFN    USHORT
ChN     USHORT
BadList ANY
NewChr  STRING(1)
    CODE
    LOOP QNdx=RECORDS(DirQ) TO 1 BY -1
        GET(DirQ,QNdx)
        DirQ:ShortName=''
        DirQ:NewName = DirQ:Name 
        LOOP X=1 TO SIZE(DirQ:Name)
            CASE VAL(DirQ:Name[X])
            OF 127 TO 255                   !127 is odd character
            OROF VAL('?')                   !04/15/20 an Unicode char is a '?'
            OROF VAL('*')                   !04/15/20 if we had an *that would be impossible
                !stop('QNdx=' & QNdx &' X=' & X &'  =' & DirQ:Name[X] )
               ChN = VAL(DirQ:Name[X])
               IF ~BadChr[ChN] THEN BadChr[ChN] = 1.
               DirQ:ShortName=CLIP(DirQ:ShortName)&CHR(ChN)
               CASE ChN
               OF 150 ; NewChr='-'  !Em Dash
               OF 160 ; NewChr='_'  !Hard Space
               OF 171 ; NewChr='['  !<<
               OF 187 ; NewChr=']'  !>>
               OF 191 ; NewChr='Q'  !Upside down ?
               OF 237 ; NewChr='i'  !i grave        Other Grave?
               ELSE     ; NewChr='`'
               END
               DirQ:NewName[X] = NewChr
            END 
       END
       LnFN = LEN( CLIP(DirQ:PathBS) & CLIP(DirQ:Name) )  !FYI path has BS
       IF LnFN > 250 THEN
          DirQ:ShortName='>250 ' & LnFN &' '& CLIP(DirQ:ShortName)
       END 
       IF DirQ:NewName = DirQ:Name AND DirQ:ShortName[1]<>'>' THEN 
          DirQ:NewName = ''
          DELETE(DirQ) ; CYCLE
       END 
       PUT(DirQ)
    END
    BadList='Bad List'
    LOOP X=80h to 255
        IF BadChr[X] THEN
            BadList=BadList &'<13,10>' & X &'  '& CHR(X)
        END
    END 
    SETCLIPBOARD(BadList)
    RETURN                 

!-----------------------------
DOO.FixProblems  PROCEDURE()
BadChr  BYTE,DIM(256) 
X       USHORT
ChN     USHORT
BadList ANY
NewChr  STRING(1)
    CODE 
    IF ~DOO.LogFileOpen() THEN RETURN False.
    LOOP QNdx=RECORDS(DirQ) TO 1 BY -1
        GET(DirQ,QNdx)
        IF ~DirQ:NewName OR DirQ:NewName = DirQ:Name THEN CYCLE. 
        RENAME(CLIP(DirQ:PathBS) & CLIP(DirQ:Name)    , |
               CLIP(DirQ:PathBS) & CLIP(DirQ:NewName) ) 
        IF ERRORCODE() THEN
            DirQ:ShortName='Err' & ErrorCode() &' '& Error()
        ELSE
            DOO.LogFileAdd()
            DirQ:ShortName='  Ok ' & DirQ:ShortName
        END 
        PUT(DirQ)
    END 
    DOO.LogFileClose()
    RETURN True
  
!-----------------------------  
DOO.LogFileOpen   PROCEDURE()!,BOOL 
ErrMsg  CSTRING(500)
    CODE
RetryLabel:    
    SHARE(LogFile) 
    CASE ErrorCode()
    OF 0 ; RETURN TRUE
    OF 2 ; CREATE(LogFile)
           IF ~ERRORCODE() THEN 
               GOTO RetryLabel:
           END
           ErrMsg='Failed to CREATE Log File?'
           DO ErrorReturn

    OF 5 ; ErrMsg='Access Denied to Log File. Is it open in Excel?'
           DO ErrorReturn
    ELSE ; ErrMsg='Failed to OPEN Log File?'
           DO ErrorReturn    
    END        
    RETURN FALSE
    
ErrorReturn ROUTINE
    Message(ErrMsg & |
           '||Error: ' & ErrorCode() &' '& ERROR() & |
            '|File: ' & ErrorFile(),'Open Log Error', ICON:Exclamation)
    RETURN FALSE 
    
DOO.LogFileAdd    PROCEDURE()  
    CODE
    CLEAR(LogFile)  
    Log:RenDate = TODAY()
    Log:RenTime = CLOCK() 
    Log:BadChar = DirQ:ShortName
    Log:Path    = DirQ:PathBS
    Log:OldName = DirQ:Name
    Log:NewName = DirQ:NewName
    ADD(LogFile)
    RETURN
DOO.LogFileClose  PROCEDURE()  
    CODE
    CLOSE(LogFile)
    RETURN 
!-----------------------------    
DOO.Debug1  PROCEDURE()    
X       USHORT
ChN     USHORT
ChrList ANY
    CODE
    GET(DirQ,CHOICE(?ListFiles))
    ChrList='' 
    !message('DOO.Debug1  PROCEDURE() TOP ') 
    LOOP X=1 to 255
        ChN = VAL(DirQ:Name[X]) 
        IF ChN = 32 THEN CYCLE.
        IF ChN < 128 THEN CYCLE.
        ChrList=ChrList &'<13,10,9>#' & X &'  ChN='& ChN &'  Txt=' & DirQ:Name[X]
    END 
    ChrList=ChrList &'<13,10><13,10>File Name: ' & CLIP(DirQ:Name) & |
                            '<13,10>File Path: ' & CLIP(DirQ:PathBS)
    SETCLIPBOARD(ChrList)
    CASE MESSAGE('The character list is on the clipboard and below:||' & ChrList,'Debug 1 - '&MsgCaption,ICON:Asterisk,'Close|Copy')
    OF 2 ; SETCLIPBOARD(ChrList)
    END 
    RETURN
!--------------------    
DB  PROCEDURE(STRING xMessage)
Prfx EQUATE('FileList: ') 
sz   CSTRING(SIZE(Prfx)+SIZE(xMessage)+1),AUTO
  CODE 
  sz  = Prfx & xMessage
  OutputDebugString( sz )