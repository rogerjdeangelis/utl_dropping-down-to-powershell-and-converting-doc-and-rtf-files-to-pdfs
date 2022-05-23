%let pgm=utl_dropping-down-to-powershell-and-converting-doc-and-rtf-files-to-pdfs;

Dropping down to powoershell and converting doc and rtf files to pdfs.

github
https://tinyurl.com/4r43zevj
https://github.com/rogerjdeangelis/utl_dropping-down-to-powershell-and-converting-doc-and-rtf-files-to-pdfs

Many Rtf to Pdf tools do not handle annotation text, like titles, footnotes, in box spanning text.
See github repo to convert pdfs to SAS datasets (two repos). This works better than R textreadr and striprtf
packages? rtf->pdf->SASA dataset.

Drop down macro below and in github
https://tinyurl.com/58pp9nz6
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

Best to convert to pdf then use my pdf 2 sas github tool.

This require ms-word.

I do not use the Limited Availability Version of MS-WORD(Office 365).
I use a perpetual verison.

If you get a pop up 'Confirm file format conversion on open', do the below.
Location: C:\Users\user name\AppData\Roaming\Microsoft\Templates.
Check to see if the "Confirm file format conversion on open" option is not active (in Word, click File >
Options > Advanced under the "General" heading).

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
You need folders
  d:/rtd
  d:/pdf
*/

ods escapechar="^";
ods rtf file="d:/rtf/class.rtf" style=minimal ;
proc report data=sashelp.class(obs=5) headskip box;
  title1 " Title 1";
  title2 " Title 2";
  footnote1 "[1] Title 1    This is a test of footnote1";
  footnote2 "[2] Title 2    This is a test of footnote2";
cols ("This is spanning ^{NEWLINE}" _all_ );
define sex / display width=5;
run;quit;
ods rtf text="^{NEWLINE}^S={font_size=8pt just=l} Bottom of boxed report";
run;quit;
ods rtf close;
title;
footnote;

/**************************************************************************************************************************/
/*                                                                                                                        */
/*  d:/rtf/class.rtf                                                                                                      */
/*                                                                                                                        */
/*                                                                                                                        */
/*  Title 1                                                                                                               */
/*  Title 2                                                                                                               */
/*                                                                                                                        */
/*   „------------------------------------------------†                                                                   */
/*   |              This is spanning                  |                                                                   */
/*   |                                                |                                                                   */
/*   |NAME      SEX          AGE     HEIGHT     WEIGHT|                                                                   */
/*   |                                                |                                                                   */
/*   ‡--------…------…----------…----------…----------‰                                                                   */
/*   |Alfred  | M    |        14|        69|     112.5|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Alice   | F    |        13|      56.5|        84|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Barbara | F    |        13|      65.3|        98|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Carol   | F    |        14|      62.8|     102.5|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Henry   | M    |        14|      63.5|     102.5|                                                                   */
/*   Š--------‹------‹----------‹----------‹----------Œ                                                                   */
/*                                                                                                                        */
/*    Bottom of boxed report                                                                                              */
/*                                                                                                                        */
/*                                                                                                                        */
/*                                                                                                                        */
/*    [1] Title 1    This is a test of footnote1                                                                          */
/*    [2] Title 2    This is a test of footnote2                                                                          */
/*                                                                                                                        */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

/**************************************************************************************************************************/
/*                                                                                                                        */
/*  d:/txt/class.txt                                                                                                      */
/*                                                                                                                        */
/*                                                                                                                        */
/*  Title 1                                                                                                               */
/*  Title 2                                                                                                               */
/*                                                                                                                        */
/*   „------------------------------------------------†                                                                   */
/*   |              This is spanning                  |                                                                   */
/*   |                                                |                                                                   */
/*   |NAME      SEX          AGE     HEIGHT     WEIGHT|                                                                   */
/*   |                                                |                                                                   */
/*   ‡--------…------…----------…----------…----------‰                                                                   */
/*   |Alfred  | M    |        14|        69|     112.5|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Alice   | F    |        13|      56.5|        84|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Barbara | F    |        13|      65.3|        98|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Carol   | F    |        14|      62.8|     102.5|                                                                   */
/*   ‡--------ˆ------ˆ----------ˆ----------ˆ----------‰                                                                   */
/*   |Henry   | M    |        14|      63.5|     102.5|                                                                   */
/*   Š--------‹------‹----------‹----------‹----------Œ                                                                   */
/*                                                                                                                        */
/*    Bottom of boxed report                                                                                              */
/*                                                                                                                        */
/*                                                                                                                        */
/*                                                                                                                        */
/*    [1] Title 1    This is a test of footnote1                                                                          */
/*    [2] Title 2    This is a test of footnote2                                                                          */
/*                                                                                                                        */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%let _inpRtf=d:/rtf/class.rtf;
%let _outPdf=d:/pdf/class.pdf;

%utlfkil(d:/pdf/class.pdf); /* this is required replace option not specified */

%utl_submit_ps64("
$word_app = New-Object -ComObject Word.Application;
    $document = $word_app.Documents.Open('&_inpRtf');
    $document.SaveAs([ref] '&_outPdf', [ref] 17);
    $document.Close();
$word_app.Quit();
");

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

805  %let _inpRtf=d:/rtf/class.rtf;  /* you can wrap a macro around the dropdown below use forward slashes*/
1806  %let _outPdf=d:/pdf/class.pdf;  /* you can wrap a macro around the dropdown below use forward slashes*/
1807  %utlfkil(d:/pdf/class.pdf); /* this is required replace option not specified */
MLOGIC(UTLFKIL):  Beginning execution.
MLOGIC(UTLFKIL):  Parameter UTLFKIL has value d:/pdf/class.pdf
MLOGIC(UTLFKIL):  %LOCAL  URC
MLOGIC(UTLFKIL):  %LET (variable name is URC)
SYMBOLGEN:  Macro variable UTLFKIL resolves to d:/pdf/class.pdf
SYMBOLGEN:  Macro variable URC resolves to 0
SYMBOLGEN:  Macro variable FNAME resolves to #LN00448
MLOGIC(UTLFKIL):  %IF condition &urc = 0 and %sysfunc(fexist(&fname)) is TRUE
MLOGIC(UTLFKIL):  %LET (variable name is URC)
SYMBOLGEN:  Macro variable FNAME resolves to #LN00448
MLOGIC(UTLFKIL):  %LET (variable name is URC)
MPRINT(UTLFKIL):   run;
MLOGIC(UTLFKIL):  Ending execution.
1808  %utl_submit_ps64("
MLOGIC(UTL_SUBMIT_PS64):  Beginning execution.
1809  $word_app = New-Object -ComObject Word.Application;
1810      $document = $word_app.Documents.Open('&_inpRtf');
SYMBOLGEN:  Macro variable _INPRTF resolves to d:/rtf/class.rtf
1811      $document.SaveAs([ref] '&_outPdf', [ref] 17);
SYMBOLGEN:  Macro variable _OUTPDF resolves to d:/pdf/class.pdf
1812      $document.Close();
1813  $word_app.Quit();
1814  ");
MLOGIC(UTL_SUBMIT_PS64):  Parameter PGM has value "$word_app = New-Object -ComObject Word.Application;    $document = $word_app.Documents.Open('d:/rtf/class.rtf');
      $document.SaveAs([ref] 'd:/pdf/class.pdf', [ref] 17);    $document.Close();$word_app.Quit();"
MLOGIC(UTL_SUBMIT_PS64):  Parameter RETURN has value
MPRINT(UTL_SUBMIT_PS64):   * write the program to a temporary file;
MPRINT(UTL_SUBMIT_PS64):   filename py_pgm "f:\wrk\_TD16084_T7610_/py_pgm.ps1" lrecl=32766 recfm=v;
MPRINT(UTL_SUBMIT_PS64):   data _null_;
MPRINT(UTL_SUBMIT_PS64):   length pgm $32755 cmd $1024;
MPRINT(UTL_SUBMIT_PS64):   file py_pgm ;
SYMBOLGEN:  Macro variable PGM resolves to "$word_app = New-Object -ComObject Word.Application;    $document = $word_app.Documents.Open('d:/rtf/class.rtf');
            $document.SaveAs([ref] 'd:/pdf/class.pdf', [ref] 17);    $document.Close();$word_app.Quit();"
MPRINT(UTL_SUBMIT_PS64):   pgm="$word_app = New-Object -ComObject Word.Application;    $document = $word_app.Documents.Open('d:/rtf/class.rtf');    $document.SaveAs([ref]
'd:/pdf/class.pdf', [ref] 17);    $document.Close();$word_app.Quit();";
MPRINT(UTL_SUBMIT_PS64):   semi=countc(pgm,';');
MPRINT(UTL_SUBMIT_PS64):   do idx=1 to semi;
MPRINT(UTL_SUBMIT_PS64):   cmd=cats(scan(pgm,idx,';'));
MPRINT(UTL_SUBMIT_PS64):   if cmd=:'. ' then cmd=trim(substr(cmd,2));
MPRINT(UTL_SUBMIT_PS64):   put cmd $char384.;
MPRINT(UTL_SUBMIT_PS64):   putlog cmd $char384.;
MPRINT(UTL_SUBMIT_PS64):   end;
MPRINT(UTL_SUBMIT_PS64):   run;

NOTE: The file PY_PGM is:
      Filename=f:\wrk\_TD16084_T7610_\py_pgm.ps1,
      RECFM=V,LRECL=32766,File Size (bytes)=0,
      Last Modified=22May2022:09:50:45,
      Create Time=22May2022:09:38:14

$word_app = New-Object -ComObject Word.Application
$document = $word_app.Documents.Open('d:/rtf/class.rtf')
$document.SaveAs([ref] 'd:/pdf/class.pdf', [ref] 17)
$document.Close()
$word_app.Quit()

NOTE: 5 records were written to the file PY_PGM.
      The minimum record length was 384.
      The maximum record length was 384.
NOTE: DATA statement used (Total process time):
      real time           0.02 seconds
      user cpu time       0.00 seconds
      system cpu time     0.01 seconds
      memory              451.65k
      OS Memory           18420.00k
      Timestamp           05/22/2022 09:50:45 AM
      Step Count                        299  Switch Count  0


MPRINT(UTL_SUBMIT_PS64):  quit;
MLOGIC(UTL_SUBMIT_PS64):  %LET (variable name is _LOC)
MLOGIC(UTL_SUBMIT_PS64):  %PUT &_loc
SYMBOLGEN:  Macro variable _LOC resolves to f:\wrk\_TD16084_T7610_\py_pgm.ps1
f:\wrk\_TD16084_T7610_\py_pgm.ps1
SYMBOLGEN:  Macro variable _LOC resolves to f:\wrk\_TD16084_T7610_\py_pgm.ps1
MPRINT(UTL_SUBMIT_PS64):   filename rut pipe "powershell.exe -executionpolicy bypass -file f:\wrk\_TD16084_T7610_\py_pgm.ps1 ";
MPRINT(UTL_SUBMIT_PS64):   data _null_;
MPRINT(UTL_SUBMIT_PS64):   file print;
MPRINT(UTL_SUBMIT_PS64):   infile rut;
MPRINT(UTL_SUBMIT_PS64):   input;
MPRINT(UTL_SUBMIT_PS64):   put _infile_;
MPRINT(UTL_SUBMIT_PS64):   putlog _infile_;
MPRINT(UTL_SUBMIT_PS64):   run;

NOTE: The infile RUT is:
      Unnamed Pipe Access Device,
      PROCESS=powershell.exe -executionpolicy bypass -file f:\wrk\_TD16084_T7610_\py_pgm.ps1,
      RECFM=V,LRECL=384

NOTE: 0 lines were written to file PRINT.
NOTE: 0 records were read from the infile RUT.
NOTE: DATA statement used (Total process time):
      real time           2.87 seconds
      user cpu time       0.00 seconds
      system cpu time     0.10 seconds
      memory              315.43k
      OS Memory           18420.00k
      Timestamp           05/22/2022 09:50:48 AM
      Step Count                        300  Switch Count  0


MPRINT(UTL_SUBMIT_PS64):   filename rut clear;
NOTE: Fileref RUT has been deassigned.
MPRINT(UTL_SUBMIT_PS64):   filename py_pgm clear;
NOTE: Fileref PY_PGM has been deassigned.
MPRINT(UTL_SUBMIT_PS64):   * use the clipboard to create macro variable;
SYMBOLGEN:  Macro variable RETURN resolves to
MLOGIC(UTL_SUBMIT_PS64):  %IF condition "&return" ^= "" is FALSE
MLOGIC(UTL_SUBMIT_PS64):  Ending execution.


/*
 _ __ ___   __ _  ___ _ __ ___
| `_ ` _ \ / _` |/ __| `__/ _ \
| | | | | | (_| | (__| | | (_) |
|_| |_| |_|\__,_|\___|_|  \___/

*/


%macro utl_submit_ps64(
      pgm
     ,return=  /* name for the macro variable from Powershell */
     )/des="Semi colon separated set of powershell commands - drop down to poershell";

  * write the program to a temporary file;
  filename py_pgm "%sysfunc(pathname(work))/py_pgm.ps1" lrecl=32766 recfm=v;
  data _null_;
    length pgm  $32755 cmd $1024;
    file py_pgm ;
    pgm=&pgm;
    semi=countc(pgm,';');
      do idx=1 to semi;
        cmd=cats(scan(pgm,idx,';'));
        if cmd=:'. ' then
           cmd=trim(substr(cmd,2));
         put cmd $char384.;
         putlog cmd $char384.;
      end;
  run;quit;
  %let _loc=%sysfunc(pathname(py_pgm));
  %put &_loc;
  filename rut pipe  "powershell.exe -executionpolicy bypass -file &_loc ";
  data _null_;
    file print;
    infile rut;
    input;
    put _infile_;
    putlog _infile_;
  run;
  filename rut clear;
  filename py_pgm clear;

  * use the clipboard to create macro variable;
  %if "&return" ^= "" %then %do;
    filename clp clipbrd ;
    data _null_;
     length txt $200;
     infile clp;
     input;
     putlog "*******  " _infile_;
     call symputx("&return",_infile_,"G");
    run;quit;
  %end;

%mend utl_submit_ps64;
