function [t, Acc pt dta sdtav  matsstat]=SSFFPunkteStat()
%%% Flip Video als Konstant eingeben
% flipv=input('Flip_Video eingeben!          ');
% flips=input('Flip_Sensor eingeben!         ');
% offset=flipv-flips;
%%% Alternative: Die Flip_wete als Konstant zu deklarieren 
%%% Für die Video 00056
flipv=13.160;
flips=4.290; % FlipSensor 00056  T= 4.290 R=4.000, L = 4.130
 
offset=flipv-flips;
AnzAkt=input('Geben Sie die Anzahl der Aktivität ein:      ');
z = input('Anfangzeile der Aktivität:         ');
%filename= uigetfile;%%Choose the figure for the Activity of the format.fig
kst1='(row,col) indices. where p<0.05. The correlated variables. ';
kst2='The Variable in Colomn I depends highly from the variable   in the next colomn';

For  i=z:AnzAkt+z-1
    %filename= uigetfile;%%Choose the figure for the Activity of the format .fig 
    %figurelist=findobj('type','figure');
    openfig(strcat('figure',num2str(z),'.fig')); %% Open the figure i.fig 
   
    [t,Acc]=ginput; %% Choose the  breakpoint of the Curve and confirm the all 
                    %%the choosen point with OK/ENTER                
    c=[t  Acc];% Matrix of the coordinates x and y 
    
    tb=c(1:end-1,1); %%% Anfangzeit dr Bewegung in der Graphik
    tend=c(2:end,1);  %%% Endezeit der Bewegung in der Graphik 
    tb=roundn(tb,-3); %%%Rounden die Zeit nach drei Kommastelle
    tend=roundn(tend,-3);%%%Rounden die Zeit nach drei Kommastelle
    Acc_b= c(1:end-1,2); %%% Startwert Acceleration
    Acc_end=c(2:end,2); %%% Endwert Acceleration
    v_acc= Acc_end-Acc_b;%%% Variation von Beschleunigung
    
    sheet = num2str(z);%%% Die Sheetsname als Nummer, den Zugriff in Matlab als in RStudio zu erleichtern
    pt= xlswrite('D1.xls', c, sheet); %Make excel file with the chosen  point or with the Matrix c
    dbew= tend-tb;   %%% Dauer der Bewegung
    Acc=Acc(2:end);
    csd=[tb tend dbew  Acc v_acc]; %%% Die Daten in Sensorzeit
    csdv=[(tb+offset) (tend+offset) tb tend  dbew  Acc v_acc];
    csdv=roundn(csdv,-3);
 %Make excel file with the chosen point or with the Matrix c
    dta= xlswrite('D2.xls', csd, sheet);
    sdtav=xlswrite('D3.xls', csdv, sheet);
    save pt;save dta;save sdtav;
    %st=csd(:,3:5);
    cdds=csdv(:,5:7);
    nva=[{'Dauer '} {'Acc  '} {' Variation Acc '} ]; {' AccY  '} {' AccZ '} {' GyrX '} {' GyrY '} {' GyrZ'} {'  MagX'} {'   MagY'} {' MagZ'}];
    sstat = struct('Merkmale', {nva},'Min',{roundn(min(cdds),-2)},'Max',{roundn(max(cdds),-2)},
        'Range',{roundn(range(cdds),-2)},'Mean',{roundn(mean(cdds),-2)},'Median',{roundn(median(cdds),-2)},…
        'Mode',{roundn(mode(cdds),-2)},'Std',{roundn(std(cdds),-2)},'Var',{roundn(var(cdds),-2)},...
       		        'IrstQ',{roundn(quantile(cdds,0.25),-2)}, 'rdQ',{roundn(quantile(cdds,.75),-2)} ,'Iqr',{roundn(iqr(cdds),-2)} ); %%%%Display die Ergebnisse
   matsstat=[sstat.Min;sstat.Max;sstat.Range;sstat.Mean;sstat.Median;
sstat.Mode; sstat.Std;sstat.Var;sstat.IrstQ; sstat.rdQ; sstat.Iqr];
   [r,p]= corrcoef(cdds);
   [ri,cj]= find(p<0.05);  % Find significant correlations.
   %%%Präsentation der Ergebnissen der statistischen Analyse
   disp('Descriptiv Statistic / Beschreibende Statistik');
   disp(sstat);disp('Descriptiv Statistic / Beschreibende Statistik');disp(cell(nva));disp(matsstat);
   disp(' Correlation coefficients R     The significant of the Correlation p-value ');
   disp([r,p]); 
            disp(strcat(kst1,kst2)); % Kommentar der  Eregebnisse

   disp([ri,cj]);% Display their (row,col) indices. where p<0.05
%%%%  Display the summary of the statistic in form of text, and array
    save sstat, save matsstat
   
    z=z+1; 
end
end





function [stfile activity]= Epstat()
    activity= 'WALK';
    
    idvideo=input('Nummer der Videos eingeben:        ');
   
    rtitle= strcat(upper('Statistische Report der Aktivität  "" _'),upper(activity), upper(' _ "" aus Video_'), num2str(idvideo));
    
    stfile=strcat('D', num2str(idvideo),'.xls');  %%% Filename  
    %stfile= uigetfile; %%% Alternative zur stfile
    csdv=xlsread(stfile,activity);
    cdds=csdv(:,5:7);
    nvsa= [{'Merkmale'},{'Min'},{'Max'},{'Range'},{'Mean'},{'Median'},{'Mode'},{'Std'},{'Var'},{'1rstQu'},{'3rdQu'},{'Iqr'}];%Statistische Variablen
    nvsa=(nvsa)' ; %%% Transpose nvsa to get it in one column
    nva=[{'Dauer '} {'AccY  '} {'Variation AccY '} ]; %%%{' AccYY  '} {'  AccYZ '} {'  AccYX '} {'   AccYY '} {'  AccYZ'} {'  MagX'} {' MagY'} {' MagZ'}];
    sstat = struct('Merkmale', {nva},'Min',{roundn(min(cdds),-2)},'Max',{roundn(max(cdds),-2)},'Range',{roundn(range(cdds),-2)},...
        'Mean',{roundn(mean(cdds),-2)},'Median',{roundn(median(cdds),-2)},'Mode',{roundn(mode(cdds),-2)},'Std',{roundn(std(cdds),-2)},'Var',{roundn(var(cdds),-2)},...
        'IrstQ',{roundn(quantile(cdds,0.25),-2)}, 'rdQ',{roundn(quantile(cdds,.75),-2)} ,'Iqr',{roundn(iqr(cdds),-2)} ); %%%%Affichage matriciel des statistiques
    matsstat=[sstat.Min;sstat.Max;sstat.Range;sstat.Mean;sstat.Median;sstat.Mode; sstat.Std;sstat.Var;sstat.IrstQ; sstat.rdQ; sstat.Iqr];
    [r,p]= corrcoef(cdds);
    [ri,cj]= find(p<0.05);  % Find significant correlations.
    %disp(sstat),disp(nva),disp(matsstat);
    %%% Visualisierung der Statistik
    disp('Descriptiv Statistic / Beschreibende Statistik');
    disp(sstat);disp('Descriptiv Statistic / Beschreibende Statistik  ');disp(cell(nva));disp(matsstat);
    disp(' Correlation coefficients R     The significant of the Correlation p-value ');
    disp([r,p]); disp('(row,col) indices. where p<0.05. The correlated variables'); disp([ri,cj]);% Display their (row,col) indices. where p<0.05
    %%%%  Display the summary of the statistic in form of text, and array
    save sstat, save matsstat
   
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    sheetb=activity;% Name der Excelblatt für die Teilergebnisse der Analyse
    k1={rtitle};
    k2= {'Matrix der Korrelationskoeffizient: Analyse der Abhängigkeit zwischen Variablen'};
    k3= {'Prüfung der Korrelation: Variablen mit höheren Korrelation Koeffizienten. Abhängigkeit zwischen Variablen'};
    k4={'P-Value: Prüfung der Validität der Korrelation zwischen Variablen [AccY.....Vbat]'};
    k5= [{'Abh.Var'};{'Unabh.Var'}] ; %%% Ab-,unabhängige Variable
    k6={'Spaltennummerierung zur Verständnis der Tabelle in Range  [B42:D43]'};
   
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Kopie von Teilergebnisse in dem Excel-Arbeitsbuch 'StatMerk{z}' 
    xlswrite('estat.xls',k1,sheetb,'C1');% Überschrift des Excelarbeitsblatts
    xlswrite('estat.xls',nva,sheetb,'B2:D2'); %%Analysierte Variablennamen
    xlswrite('estat.xls',nvsa,sheetb,'A2:A13'); %% Statistiche Merkmale in der 1.Spalten der Exceltabelle
    xlswrite('estat.xls',matsstat,sheetb,'B3:D13');%%% Werte von stistische Bewertung
    xlswrite('estat.xls',k2,sheetb,'A16');% Überschrift der Korrelationsmatrix
    xlswrite('estat.xls',k4,sheetb,'D16');
    
    xlswrite('estat.xls',nva,sheetb,'B19:D19');% Überschrift der Korrelationsmatrix
    xlswrite('estat.xls',nva,sheetb,'E19:G19');%% P-Value. Matrixkopf 
    xlswrite('estat.xls',(nva)',sheetb,'A20:A22');
    xlswrite('estat.xls',[r,p],sheetb,'B20:G22');%%% Korrelationskoefizienten
    
    xlswrite('estat.xls',k6,sheetb,'A27');
    xlswrite('estat.xls',{'Variable'},sheetb,'A29');
    xlswrite('estat.xls',nva,sheetb,'B29:D29')%%% Namen der Variablen
    xlswrite('estat.xls',{'Spaltennummer'},sheetb,'A30');
    xlswrite('estat.xls',1:3,sheetb,'B30:D30')%%% Nummerierung der Variablen
    
    
    xlswrite('estat.xls',k3,sheetb,'A32');%Überschrift für die Variablemen mit höheren Korrelation
    xlswrite('estat.xls',k5,sheetb,'A34:A35');
    xlswrite('estat.xls',([ri,cj])',sheetb,'B34:G35');%3 Variablen insgesamt und 
    %%%in Maximum kann der Wert einer Variable abhängig von den 2 anderen
    %%%Variablen sein.dh. 3Beobachtete Variable*2Anderen Variablen*2Zeile/Variable
    
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
end


% How to get the coordinate of the points of the curve of Acceleration and make directly statistics? This function must be put in the file of the ploted graphik otherwise  ,
% it will not work. The statistics are displayed in the specified excel workbooks 'sstatx.,-y.,-z.xls'

function[tx,ty tz AccX ptx dtax sdtavx AccY pty dtay sdtavy AccZ ptz dtaz sdtavz]=SPAccXYZ() %%% Record the coordinate of the inportanspoint in the three axis
%%% Flip Video als Konstant eingeben
% flipv=input('Flip_Video eingeben!          ');
% flips=input('Flip_Sensor eingeben!         ');
% offset=flipv-flips;
%%% Alternative: Die Flip_wete als Konstant zu deklarieren 
%%% FLIP_Zeit als Konstant für die Video eingeben

%%% Video 0030 Flip[V T  R L]=[12.640 10.410  10.010  10.010]
%%% Video 0056:[V  T; R; L]=    [13.160  4.290  4.000  4.130] 
%%% Video 0064 [V  T  R  L]=    [13.360  9.310  9.700  9.800]
%%% Video 0066 [V  T  R  L]=    [13.560  6.760  6.699  6.530]
%%% Video 0060 Flip[V=   9.000  T= 6.300, R= 6.299, L =6.010]
%%% Video 141:[V  T; R; L]=     [12.640; 4.950; 4.200; 4.580] 
%%% Video 111:[V  T; R; L]=     [12.160; 9.490; 9.420; 8.900] 
%%% Video 114:[V  T; R; L]=     [7.680; 4.060 ; 3.750; 3.720] 
%%% Video 0061:[V  T; R; L]=     [11.480  6.740  7.900 6.759]
%%% Video 0051:[V  T; R; L]=     [10.280  6.050  6.049 6.000]
%%% Video 137:[V  T; R; L]=      [11.040 6.000  6.000  5.950]
%%% Video 100:[V  T; R; L]=      [10.040 6.310  6.778  5.480]

flipv=10.280;
flips=6.050; % FlipSensor 00056  T=10.410 R=10.010, L = 10.010

offset=flipv-flips;
AnzAkt=input('Geben Sie die Anzahl der Aktivität ein:      ');
z = input('Anfangzeile der Aktivität:         ');
%filename= uigetfile;%%Choose the figure for the Activity of the format.fig

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Titel von Excelblatt
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% zur Presentation der
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Ergebnisse
%%%btitle=strcat('Teilergebnis:Statistische Analyse der Aktivität[',activity,'] aus Video [',num2str(idvideo),']mit Daten aus Sensor[', idsensor,']  mit Offset=[',num2str(offset),'Sec]');

prompt= {'Nummer der Video eingeben',...
         'Geben Sie das Sensor Präfix (T/R/L)ein:',...
         'Geben Sie die untersuchte Aktivität ein!'};

dlg_title='Die  Input Variablen für die Untersuchung';
inputd   =inputdlg(prompt,dlg_title);
idvideo  =inputd(1);
idsensor =upper(inputd(2));
activity =upper(inputd(3));
%%%btitle=strcat('Teilergebnis:Statistische Analyse der Aktivität[',activity,']aus Video [',num2str(idvideo),'] mit Daten aus Sensor[', idsensor,']  mit Offset=[',num2str(offset),'Sec]');
t1={'Teilergebnisse: Statistische Analyse der Aktivität'}; 
t2={' aus Video'};
t3={'mit Daten aus Sensor'};
t4={'mit Offset'};

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

for i=z:AnzAkt+z-1
    %filename= uigetfile;%%Choose the figure for the Activity of the format .fig 
    %figurelist=findobj('type','figure');
    openfig(strcat('figure',num2str(z),'.fig')); %% Open the figure i.fig 
   %findpeaks(strcat('figure',num2str(z),'.fig'),'NPeaks',15);
    [tx,AccX]=ginput; %% Choose the  breakpoint of the Curve and confirm the all 
                    %%the choosen point with OK/ENTER 
	[ty,AccY]=ginput;
	[tz,AccZ]=ginput;
    cx=[tx AccX];% Matrix of the coordinates x and y 
    cy=[ty AccY];
	cz=[tz AccZ];
    %Time and time duration of movements 
    tbx=cx(1:end-1,1); %%% Anfangzeit dr Bewegung in der Graphik
    tendx=cx(2:end,1);  %%% Endezeit der Bewegung in der Graphik 
    tbx=roundn(tbx,-3); %%%Rounden die Zeit nach drei Kommastelle
    tendx=roundn(tendx,-3);%%%Rounden die Zeit nach drei Kommastelle
	%%% Variation of Acceleration AccX
  
    Acc_bx= cx(1:end-1,2); %%% Startwert Acceleration
    Acc_endx=cx(2:end,2); %%% Endwert Acceleration
    v_accx= Acc_endx-Acc_bx;%%% Variation von Beschleunigung
    %%%
    
    %Time and time duration of movements 
	tby=cy(1:end-1,1); %%% Anfangzeit dr Bewegung in der Graphik
    tendy=cy(2:end,1);  %%% Endezeit der Bewegung in der Graphik 
    tby=roundn(tby,-3); %%%Rounden die Zeit nach drei Kommastelle
    tendy=roundn(tendy,-3);%%%Rounden die Zeit nach drei Kommastelle
	%%% Variation of Acceleration AccY
     
    Acc_by= cy(1:end-1,2); %%% Startwert Acceleration
    Acc_endy=cy(2:end,2); %%% Endwert Acceleration
    v_accy= Acc_endy-Acc_by;%%% Variation von Beschleunigung
    %%%
    
    %Time and time duration of movements 
	tbz=cz(1:end-1,1); %%% Anfangzeit dr Bewegung in der Graphik
    tendz=cz(2:end,1);  %%% Endezeit der Bewegung in der Graphik 
    tbz=roundn(tbz,-3); %%%Rounden die Zeit nach drei Kommastelle
    tendz=roundn(tendz,-3);%%%Rounden die Zeit nach drei Kommastelle
    %%% Variation of Acceleration AccZ
  
    Acc_bz= cz(1:end-1,2); %%% Startwert Acceleration
    Acc_endz=cz(2:end,2); %%% Endwert Acceleration
    v_accz= Acc_endz-Acc_bz;%%% Variation von Beschleunigung
    %%%
    
    sheet = num2str(z); %%% Die Sheetname als nummer, den Zugriff in Matlab als in RStudio zu erleichtern
    
	ptx= xlswrite('Dax1.xls', cx, sheet); %Make excel file with the chosen point or with the Matrix c
    dbewx= tendx-tbx;             %%% Dauer der Bewegung
    AccX=AccX(2:end);
   
    csdx=[tbx tendx dbewx  AccX  v_accx]; %%% Die Daten in Sensorzeit
    csdvx=[(tbx+offset) (tendx+offset) tbx tendx  dbewx  AccX v_accx];
    csdvx=roundn(csdvx,-3);
	
    pty= xlswrite('Day1.xls', cx, sheet); %Make excel file with the chosen point or with the Matrix c
    dbewy= tendy-tby;                     %%% Dauer der Bewegung
    AccY=AccY(2:end);
    csdy=[tby tendy dbewy  AccY  v_accy ];         %%% Die Daten in Sensorzeit
    csdvy=[(tby+offset) (tendy+offset) tby tendy  dbewy  AccY  v_accy]; 
    csdvy=roundn(csdvy,-3);
	
    ptz= xlswrite('Daz1.xls', cx, sheet); %Make excel file with the chosen point or with the Matrix c
    dbewz= tendz-tbz;   %%% Dauer der Bewegung
    AccZ=AccZ(2:end);
    csdz=[tbz tendz dbewz  AccZ v_accz]; %%% Die Daten in Sensorzeit
    csdvz=[(tbz+offset) (tendz+offset) tbz tendz  dbewz  AccZ v_accz];
    csdvz=roundn(csdvz,-3);
	
    %%% Save Data in Excel Data for statistics purposes
    
    dtax= xlswrite('Dax2.xls', csdx, sheet);
    sdtavx=xlswrite('Dax3.xls', csdvx, sheet);
    save ptx;save dtax;save sdtavx;
	
	dtay= xlswrite('Day2.xls', csdy, sheet);
    sdtavy=xlswrite('Day3.xls', csdvy, sheet);
    save pty;save dtay;save sdtavy;
	
	dtaz= xlswrite('Daz2.xls', csdz, sheet);
    sdtavz=xlswrite('Daz3.xls',csdvz,sheet);
    save ptz;save dtaz;save sdtavz;
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%          STATISTIK      %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %%% Selection of the importentant variables or parts of saved data for statistik
    %%% Reading die Data from the files written above. 
    rsdx=xlsread('Dax3.xls',sheet);
    rsdy=xlsread('Day3.xls',sheet);
    rsdz=xlsread('Daz3.xls',sheet);
    cddsx=rsdx(:,5:7); %%%[ dbewx  AccX  v_accx];
    cddsy=rsdy(:,5:7); %%%[ dbewy  AccY  v_accy];
    cddsz=rsdz(:,5:7); %%%[ dbewz  AccZ  v_accZ];
    
    nvsa= [{'Merkmale'},{'Min'},{'Max'},{'Range'},{'Mean'},{'Median'},{'Mode'},{'Std'},{'Var'},{'1rstQu'},{'3rdQu'},{'Iqr'}];%Statistische Variablen
    nvsa=(nvsa)' ; %%% Transpose nvsa to get it in one column
    nvax=[{'Dauer '} {'AccX  '} {'Variation AccX'} ]; %%%{' AccY  '} {'  AccZ '} {'  GyrX '} {'   GyrY '} {'  GyrZ'} {'  MagX'} {' MagY'} {' MagZ'}];
    nvay=[{'Dauer '} {'AccY  '} {'Variation AccY'} ];
    nvaz=[{'Dauer '} {'AccZ  '} {'Variation AccZ'} ];
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Berechnung der
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% statistischen Merkmale
    %%% Für AccX
    sstatx = struct('Merkmale', {nvax},'Min',{roundn(min(cddsx),-3)},'Max',{roundn(max(cddsx),-3)},'Range',{roundn(range(cddsx),-3)},...
        'Mean',{roundn(mean(cddsx),-3)},'Median',{roundn(median(cddsx),-3)},'Mode',{roundn(mode(cddsx),-3)},'Std',{roundn(std(cddsx),-3)},'Var',{roundn(var(cddsx),-3)},...
        'IrstQ',{roundn(quantile(cddsx,0.25),-3)}, 'rdQ',{roundn(quantile(cddsx,.75),-3)} ,'Iqr',{roundn(iqr(cddsx),-3)} ); %%%%Affichage matriciel des statistiques
    matsstatx=[sstatx.Min;sstatx.Max;sstatx.Range;sstatx.Mean;sstatx.Median;sstatx.Mode; sstatx.Std;sstatx.Var;sstatx.IrstQ; sstatx.rdQ; sstatx.Iqr];
    [r_x,p_x]= corrcoef(cddsx);
    [ri_x,cj_x]= find(p_x<0.05);  % Find significant correlations.
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%% Für AccY
    sstaty = struct('Merkmale', {nvay},'Min',{roundn(min(cddsy),-3)},'Max',{roundn(max(cddsy),-3)},'Range',{roundn(range(cddsy),-3)},...
        'Mean',{roundn(mean(cddsy),-3)},'Median',{roundn(median(cddsy),-3)},'Mode',{roundn(mode(cddsy),-3)},'Std',{roundn(std(cddsy),-3)},'Var',{roundn(var(cddsy),-3)},...
        'IrstQ',{roundn(quantile(cddsy,0.25),-3)}, 'rdQ',{roundn(quantile(cddsy,.75),-3)} ,'Iqr',{roundn(iqr(cddsy),-3)} ); %%%%Affichage matriciel des statistiques
    matsstaty=[sstaty.Min;sstaty.Max;sstaty.Range;sstaty.Mean;sstaty.Median;sstaty.Mode; sstaty.Std;sstaty.Var;sstaty.IrstQ; sstaty.rdQ; sstaty.Iqr];
    [r_y,p_y]= corrcoef(cddsy);%%% Matrix der Korrelationskoefizienten und P-Value 
    [ri_y,cj_y]= find(p_y<0.05);  % Find significant correlations.
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Für AccZ
    sstatz = struct('Merkmale', {nvaz},'Min',{roundn(min(cddsz),-3)},'Max',{roundn(max(cddsz),-3)},'Range',{roundn(range(cddsz),-3)},...
                   'Mean',{roundn(mean(cddsz),-3)},'Median',{roundn(median(cddsz),-3)},'Mode',{roundn(mode(cddsz),-3)},'Std',{roundn(std(cddsz),-3)},'Var',{roundn(var(cddsz),-3)},...
                   'IrstQ',{roundn(quantile(cddsz,0.25),-3)}, 'rdQ',{roundn(quantile(cddsz,.75),-3)} ,'Iqr',{roundn(iqr(cddsz),-3)} ); %%%%Affichage matriciel des statistiques
    matsstatz=[sstatz.Min;sstatz.Max;sstatz.Range;sstatz.Mean;sstatz.Median;sstatz.Mode; sstatz.Std;sstatz.Var;sstatz.IrstQ; sstatz.rdQ; sstatz.Iqr];
    [r_z,p_z]= corrcoef(cddsz);%%% Matrix der Korrelationskoefizienten und P-Value 
    [ri_z,cj_z]= find(p_z<0.05);  % Find significant correlations. Validität der Korrelation bis einem Konfidenzintervall bis grösser gleich 0.95
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Tiel von Tabellen   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    sheetb=strcat('StMerk',num2str(z));% Name der Excelblatt für die Teilergebnisse der Analyse
    k2= {'Matrix der Korrelationskoeffizient: Analyse der Abhängigkeit zwischen Variablen'};
    k3= {'Prüfung der Korrelation: Variablen mit höheren Korrelation Koeffizienten. Abhängigkeit zwischen Variablen'};
    k4= {'P-Value: Prüfung der Validität der Korrelation zwischen Variablen [AccX.....Vbat]'};
    k5= [{'Abh.Var'};{'Unabh.Var'}] ; %%% Ab-,unabhängige Variable
    k6= {'Spaltennummerierung zur Verständnis der Tabelle in Range  [B42:D43]'};
   
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  STATISTIK PRÄNSENTATION_Y  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%% Entitlung der Excelblatt für AccX
    xlswrite('sstatx.xls',t1,sheetb,'A1');
    xlswrite('sstatx.xls',activity,sheetb,'E1');
    xlswrite('sstatx.xls',t2,sheetb,'F1');
    xlswrite('sstatx.xls',idvideo,sheetb,'G1');
    xlswrite('sstatx.xls',t3,sheetb,'H1');
    xlswrite('sstatx.xls',idsensor,sheetb,'J1');
    xlswrite('sstatx.xls',t4,sheetb,'K1');
    xlswrite('sstatx.xls',offset,sheetb,'L1');
    %xlswrite('sstatx.xls',t5,sheetb,'M1');

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Kopie von Teilergebnisse in dem Excel-Arbeitsbuch 'StatMerk{z}' 
    %xlswrite('sstat.xls',k1,sheetb,'B1');% Überschrift des Excelarbeitsblatts 
    xlswrite('sstatx.xls',nvax,sheetb,'B2:D2'); %%Analysierte Variablennamen
    xlswrite('sstatx.xls',nvsa,sheetb,'A2:A13'); %% Statistiche Merkmale in der 1.Spalten der Exceltabelle
    xlswrite('sstatx.xls',matsstatx,sheetb,'B3:D13');%%% Werte von stistische Bewertung
    xlswrite('sstatx.xls',k2,sheetb,'A16');% Überschrift der Korrelationsmatrix
    xlswrite('sstatx.xls',k4,sheetb,'D16');
    
    xlswrite('sstatx.xls',nvax,sheetb,'B19:D19');% Überschrift der Korrelationsmatrix
    xlswrite('sstatx.xls',nvax,sheetb,'E19:G19');%% P-Value. Matrixkopf 
    xlswrite('sstatx.xls',(nvax)',sheetb,'A20:A22');
    xlswrite('sstatx.xls',[r_x,p_x],sheetb,'B20:G22');%%% Korrelationskoefizienten
    
    xlswrite('sstatx.xls',k6,sheetb,'A27');
    xlswrite('sstatx.xls',{'Variable'},sheetb,'A29');
    xlswrite('sstatx.xls',nvax,sheetb,'B29:D29')%%% Namen der Variablen
    xlswrite('sstatx.xls',{'Spaltennummer'},sheetb,'A30');
    xlswrite('sstatx.xls',1:3,sheetb,'B30:D30')%%% Nummerierung der Variablen
    
    
    xlswrite('sstatx.xls',k3,sheetb,'A32');%Überschrift für die Variablemen mit höheren Korrelation
    xlswrite('sstatx.xls',k5,sheetb,'A34:A35');
    xlswrite('sstatx.xls',([ri_x,cj_x])',sheetb,'B34:D35');%3 Variablen insgesamt und 
    
    
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% STATISTIK PRÄNSENTATION_Y %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%% Entitlung der Excelblatt für AccY
    xlswrite('sstaty.xls',t1,sheetb,'A1');
    xlswrite('sstaty.xls',activity,sheetb,'E1');
    xlswrite('sstaty.xls',t2,sheetb,'F1');
    xlswrite('sstaty.xls',idvideo,sheetb,'G1');
    xlswrite('sstaty.xls',t3,sheetb,'H1');
    xlswrite('sstaty.xls',idsensor,sheetb,'J1');
    xlswrite('sstaty.xls',t4,sheetb,'K1');
    xlswrite('sstaty.xls',offset,sheetb,'L1');
    %xlswrite('sstaty.xls',t5,sheetb,'M1');
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Kopie von Teilergebnisse in dem Excel-Arbeitsbuch 'StatMerk{z}' 
    %xlswrite('sstat.xls',k1,sheetb,'B1');% Überschrift des Excelarbeitsblatts
    xlswrite('sstaty.xls',nvay,sheetb,'B2:D2'); %%Analysierte Variablennamen
    xlswrite('sstaty.xls',nvsa,sheetb,'A2:A13'); %% Statistiche Merkmale in der 1.Spalten der Exceltabelle
    xlswrite('sstaty.xls',matsstaty,sheetb,'B3:D13');%%% Werte von stistische Bewertung
    xlswrite('sstaty.xls',k2,sheetb,'A16');% Überschrift der Korrelationsmatrix
    xlswrite('sstaty.xls',k4,sheetb,'D16');
    
    xlswrite('sstaty.xls',nvay,sheetb,'B19:D19');% Überschrift der Korrelationsmatrix
    xlswrite('sstaty.xls',nvay,sheetb,'E19:G19');%% P-Value. Matrixkopf 
    xlswrite('sstaty.xls',(nvay)',sheetb,'A20:A22');
    xlswrite('sstaty.xls',[r_y,p_y],sheetb,'B20:G22');%%% Korrelationskoefizienten
    
    xlswrite('sstaty.xls',k6,sheetb,'A27');
    xlswrite('sstaty.xls',{'Variable'},sheetb,'A29');
    xlswrite('sstaty.xls',nvay,sheetb,'B29:D29')%%% Namen der Variablen
    xlswrite('sstaty.xls',{'Spaltennummer'},sheetb,'A30');
    xlswrite('sstaty.xls',1:3,sheetb,'B30:D30')%%% Nummerierung der Variablen
    
    
    xlswrite('sstaty.xls',k3,sheetb,'A32');%Überschrift für die Variablemen mit höheren Korrelation
    xlswrite('sstaty.xls',k5,sheetb,'A34:A35');
    xlswrite('sstaty.xls',([ri_y,cj_y])',sheetb,'B34:D35');%3 Variablen insgesamt und 
    
    
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   STATISTIK PRÄNSENTATION_Z   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Entitlung der Excelblatt für AccZ
    xlswrite('sstatz.xls',t1,sheetb,'A1');
    xlswrite('sstatz.xls',activity,sheetb,'E1');
    xlswrite('sstatz.xls',t2,sheetb,'F1');
    xlswrite('sstatz.xls',idvideo,sheetb,'G1');
    xlswrite('sstatz.xls',t3,sheetb,'H1');
    xlswrite('sstatz.xls',idsensor,sheetb,'J1');
    xlswrite('sstatz.xls',t4,sheetb,'K1');
    xlswrite('sstatz.xls',offset,sheetb,'L1');
    %xlswrite('sstatz.xls',t5,sheetb,'M1'); 
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Kopie von Teilergebnisse in dem Excel-Arbeitsbuch 'StatMerk{z}' 
    %xlswrite('sstat.xls',k1,sheetb,'B1');% Überschrift des Excelarbeitsblatts
    xlswrite('sstatz.xls',nvaz,sheetb,'B2:D2'); %%Analysierte Variablennamen
    xlswrite('sstatz.xls',nvsa,sheetb,'A2:A13'); %% Statistiche Merkmale in der 1.Spalten der Exceltabelle
    xlswrite('sstatz.xls',matsstatz,sheetb,'B3:D13');%%% Werte von stistische Bewertung
    xlswrite('sstatz.xls',k2,sheetb,'A16');% Überschrift der Korrelationsmatrix
    xlswrite('sstatz.xls',k4,sheetb,'D16');
    
    xlswrite('sstatz.xls',nvaz,sheetb,'B19:D19');% Überschrift der Korrelationsmatrix
    xlswrite('sstatz.xls',nvaz,sheetb,'E19:G19');%% P-Value. Matrixkopf 
    xlswrite('sstatz.xls',(nvaz)',sheetb,'A20:A22');
    xlswrite('sstatz.xls',[r_z,p_z],sheetb,'B20:G22');%%% Korrelationskoefizienten
    
    xlswrite('sstatz.xls',k6,sheetb,'A27');
    xlswrite('sstatz.xls',{'Variable'},sheetb,'A29');
    xlswrite('sstatz.xls',nvaz,sheetb,'B29:D29')%%% Namen der Variablen
    xlswrite('sstatz.xls',{'Spaltennummer'},sheetb,'A30');
    xlswrite('sstatz.xls',1:3,sheetb,'B30:D30')%%% Nummerierung der Variablen
    
    
    xlswrite('sstatz.xls',k3,sheetb,'A32');%Überschrift für die Variablemen mit höheren Korrelation
    xlswrite('sstatz.xls',k5,sheetb,'A34:A35');
    xlswrite('sstatz.xls',([ri_z,cj_z])',sheetb,'B34:D35');%3 Variablen insgesamt und 
      
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% STATISTIK PRÄNSENTATION_Z%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
    z=z+1; 
end
end



function[m]=erhmean()
activity= 'MT';

z= input('Anzahl der Videos oder Probanden eingeben:       ');
ii=1;
for i=ii:z;
    idvideo=input('ID_Video oder Probanden  eingeben:       ');
    stfile= strcat('E',num2str(idvideo),'.xls');
    %bi=cellstr(strcat('B',num2str(ii),':D',num2str(ii)));
    m= xlsread(stfile,activity,'B6:D6');
    xlswrite(activity, m, num2str(idvideo));
    ii=ii+1;
end
end


% Combdst()

% Co



function[hi ci stats groupnames  hiii piii ciii]=fanov()

activity='GRAP';
ext='.xls';
%nva  = [{'Dauer'} {'AccY'} {'Var_AccY'}]; % For activity='WALK'
nva= [{'Dauer'} {'GyrX'} {'Var_GyrX'}]; % For activity='PICK';'GRAP'; 'MT'
%nva= [{'Dauer'} {'GyrZ'} {'Var_GyrZ'}]; % For activity='PLACE';'PUSH'
sfile=strcat(activity,ext);
sheet=strcat('G',activity);
df=xlsread(sfile,sheet);
groupnames=num2str(df(:,1));
cniveau= input('Die gewünschte Signifikanzniveau[0 1]for the t_test eingeben:  ');
[hi ci stats]=anova1((df(:,2:4))',groupnames);
multcompare(stats);

[hiii piii ciii] = ttest(df(:,2:4),0, cniveau);

s1= strcat(activity, '_fm'); %
s2= strcat(activity,'_bpm');

mkdir(s1); %
mkdir(s2); %

figurelist=findobj('type','figure');
for j=1:numel(figurelist)
    saveas(figurelist(j),fullfile(s1,['figure' num2str(figurelist(j)) '.fig']));%%%% speichert in Format fig
    saveas(figurelist(j),fullfile(s2,['figure' num2str(figurelist(j)) '.bmp']));%%%% speichert in Format .bmp
end
close;

sheet=strcat('TS',activity);%%% Sheet for recording t_test value
tfile=sfile; % This file will contain the t_statistics
%[hiii  piii ciii]=ttest(sfs,0, cniveau);

xlswrite(tfile,{upper(strcat('t-test  for  sample analysis _',activity))},sheet,'A1');
xlswrite(tfile, nva, sheet,'B4:D4');% Datenvariable
xlswrite(tfile, {'HYPOTHESIS'},sheet, 'A5');
xlswrite(tfile,hiii, sheet, 'B5:D5');

xlswrite(tfile, nva, sheet, 'B7:D7');% Datenvariable for P-Value
xlswrite(tfile,{'PROBABILITY'}, sheet, 'A8');
xlswrite(tfile,piii, sheet, 'B8:D8'); % p_wert nach test von Student
st1=upper('Konfidenzintervall mit Signifikanzniveau     " " ');
stt=cellstr(strcat(st1,num2str(1-cniveau), ' " "'));
xlswrite(tfile,stt, sheet, 'A10'); %%% Titel
xlswrite(tfile, nva, sheet, 'B12:D12');% Datenvariable
xlswrite(tfile,{'LOWER_BOUND'; 'UPPER_BOUND'},sheet,'A13:A14');
xlswrite(tfile, ciii, sheet, 'B13:D14');% Value von confidence Intervall 
xlswrite(tfile, {'UNTERSUCHTE VIDEOS FÜR AKTIVITÄT _'}, sheet, 'A16');
xlswrite(tfile, {activity}, sheet, 'D16');
xlswrite(tfile, (df(:,1))', sheet, 'A18');
end
