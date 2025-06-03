# utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell
Importing a sas created rtf file into ms word using powershell
    %let pgm=utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell;

    %stop-submission;

    Importing a sas created rtf file into ms word using powershell

    github
    https://tinyurl.com/2ps7jk5t
    https://github.com/rogerjdeangelis/utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell

    sas communities
    https://tinyurl.com/ye3ske6e
    https://communities.sas.com/t5/SAS-Programming/How-to-close-word-file-programmatically-from-SAS/m-p/790074#M252926


    ms doc output
    https://tinyurl.com/yrn65hzs
    https://github.com/rogerjdeangelis/utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell/blob/main/rtf2ms.docx

    rtf input
    https://tinyurl.com/2yju37pu
    https://github.com/rogerjdeangelis/utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell/blob/main/rtf2ms.rtf

    peplexity ai query
    https://www.perplexity.ai/search/how-can-copy-an-rtf-file-to-mi-jBZbff6sQVG0W4H69YGoEw
    how can copy an rtf file to microsoft word using powershell


    Related
    ------------------------------------------------------------------------------------------------------------------------------------
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-return-the-number-of-lines-in-a-file-to-wps
    https://github.com/rogerjdeangelis/utl-drop-down-using-dosubl-from-sas-datastep-to-wps-r-perl-powershell-python-msr-vb
    https://github.com/rogerjdeangelis/utl-examples-of-drop-downs-from-sas-to-wps-r-microsoftR-python-perl-powershell
    https://github.com/rogerjdeangelis/utl-how-to-remove-a-record-from-a-csv-file-in-place-with-powershell-fast
    https://github.com/rogerjdeangelis/utl-powershell-unzip-one-meber-of-a-winzip-archive
    https://github.com/rogerjdeangelis/utl-read-print-file-backwards-in-perl-powershell-sas-r-and-python
    https://github.com/rogerjdeangelis/utl-scraping-AI-results-without-restriction-or-API-with-powershell-and-perplexity
    https://github.com/rogerjdeangelis/utl_dropping-down-to-powershell-and-converting-doc-and-rtf-files-to-pdfs
    https://github.com/rogerjdeangelis/zip-and-unzip-using-ms-powershell


    /*****************************************************************************************************************************/
    /*                                      |                                               |                                    */
    /*             INPUT                    |             PROCESS                           |           OUTPUT                   */
    /*             =====                    |             =======                           |            ======                  */
    /* https://tinyurl.com/2yju37pu         |                                               |  https://tinyurl.com/yrn65hzs      */
    /* d:/rtf/rtf2ms.rtf                    |                                               |  d:/doc/rtf2ms.docx                */
    /* -------------------------------      | %utlfkil(d:/doc/rtf2ms.docx);                 |  -------------------------------   */
    /* |   |           Sex           |      |                                               |  |   |           Sex           |   */
    /* |   |-------------------------|      | %utl_psbegin;                                 |  |   |-------------------------|   */
    /* |   |     F      |     M      |      | parmcards4;                                   |  |   |     F      |     M      |   */
    /* |   |------------+------------|      | # Create a Word application COM object        |  |   |------------+------------|   */
    /* |   |  N  | PctN |  N  | PctN |      |  $word=New-Object -ComObject Word.Application |  |   |  N  | PctN |  N  | PctN |   */
    /* |---+-----+------+-----+------|      | $word.Visible = $false                        |  |---+-----+------+-----+------|   */
    /* |Age|     |      |     |      |      |                                               |  |Age|     |      |     |      |   */
    /* |---|     |      |     |      |      | # Path to the source RTF file                 |  |---|     |      |     |      |   */
    /* |11 | 1.00| 11.11| 1.00| 10.00|      | $rtfPath = "d:/rtf/rtf2ms.rtf"                |  |11 | 1.00| 11.11| 1.00| 10.00|   */
    /* |---+-----+------+-----+------|      |                                               |  |---+-----+------+-----+------|   */
    /* |12 | 2.00| 22.22| 3.00| 30.00|      | # Path to save the Word document              |  |12 | 2.00| 22.22| 3.00| 30.00|   */
    /* |---+-----+------+-----+------|      | $docxPath = "d:/doc/rtf2ms.docx"              |  |---+-----+------+-----+------|   */
    /* |13 | 2.00| 22.22| 1.00| 10.00|      |                                               |  |13 | 2.00| 22.22| 1.00| 10.00|   */
    /* |---+-----+------+-----+------|      | # Open the RTF document                       |  |---+-----+------+-----+------|   */
    /* |14 | 2.00| 22.22| 2.00| 20.00|      | $doc = $word.Documents.Open($rtfPath)         |  |14 | 2.00| 22.22| 2.00| 20.00|   */
    /* |---+-----+------+-----+------|      |                                               |  |---+-----+------+-----+------|   */
    /* |15 | 2.00| 22.22| 2.00| 20.00|      | # Save as (wdFormatXMLDocument=12 docx)       |  |15 | 2.00| 22.22| 2.00| 20.00|   */
    /* |---+-----+------+-----+------|      | $wdFormatXMLDocument = 12                     |  |---+-----+------+-----+------|   */
    /* |16 |    .|     .| 1.00| 10.00|      | $doc.SaveAs([ref] $docxPath                   |  |16 |    .|     .| 1.00| 10.00|   */
    /* -------------------------------      |  ,[ref] $wdFormatXMLDocument)                 |  -------------------------------   */
    /*                                      | # Close the document and quit Word            |                                    */
    /* ods rtf file="d:/rtf/rtf2ms.rtf";    | $doc.Close()                                  |                                    */
    /*                                      | $word.Quit()                                  |                                    */
    /* proc tabulate data=sashelp.class;    | ;;;;                                          |                                    */
    /*  class sex age;                      | %utl_psend;                                   |                                    */
    /*  table age,sex*(n pctn<age>)/rts=8;  |                                               |                                    */
    /* run;quit;                            |                                               |                                    */
    /*                                      |                                               |                                    */
    /* ods rtf close;                       |                                               |                                    */
    /*****************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */
    /*---- for development purposes ----*/
    %utl_close;
    %utlfkil(d:/rtf/rtf2ms.rtf);

    ods rtf file="d:/rtf/rtf2ms.rtf";

    proc tabulate data=sashelp.class;
     class sex age;
     table age,sex*(n pctn<age>)/rts=8;
    run;quit;

    ods rtf close;

    /**************************************************************************************************************************/
    /* https://tinyurl.com/2yju37pu                                                                                           */
    /* d:/rtf/rtf2ms.rtf                                                                                                      */
    /* -------------------------------                                                                                        */
    /* |   |           Sex           |                                                                                        */
    /* |   |-------------------------|                                                                                        */
    /* |   |     F      |     M      |                                                                                        */
    /* |   |------------+------------|                                                                                        */
    /* |   |  N  | PctN |  N  | PctN |                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |Age|     |      |     |      |                                                                                        */
    /* |---|     |      |     |      |                                                                                        */
    /* |11 | 1.00| 11.11| 1.00| 10.00|                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |12 | 2.00| 22.22| 3.00| 30.00|                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |13 | 2.00| 22.22| 1.00| 10.00|                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |14 | 2.00| 22.22| 2.00| 20.00|                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |15 | 2.00| 22.22| 2.00| 20.00|                                                                                        */
    /* |---+-----+------+-----+------|                                                                                        */
    /* |16 |    .|     .| 1.00| 10.00|                                                                                        */
    /* -------------------------------                                                                                        */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utlfkil(d:/doc/rtf2ms.docx);

    %utl_psbegin;
    parmcards4;
    # Create a Word application COM object
    $word=New-Object -ComObject Word.Application
    $word.Visible = $false

    # Path to the source RTF file
    $rtfPath = "d:/rtf/rtf2ms.rtf"

    # Path to save the Word document
    $docxPath = "d:/doc/rtf2ms.docx"

    # Open the RTF document
    $doc = $word.Documents.Open($rtfPath)

    # Save as (wdFormatXMLDocument=12 docx)
    $wdFormatXMLDocument = 12
    $doc.SaveAs([ref] $docxPath
     ,[ref] $wdFormatXMLDocument)
    # Close the document and quit Word
    $doc.Close()
    $word.Quit()
    ;;;;
    %utl_psend;

    /**************************************************************************************************************************/
    /*   https://tinyurl.com/yrn65hzs                                                                                         */
    /*   d:/doc/rtf2ms.docx                                                                                                   */
    /*   -------------------------------                                                                                      */
    /*   |   |           Sex           |                                                                                      */
    /*   |   |-------------------------|                                                                                      */
    /*   |   |     F      |     M      |                                                                                      */
    /*   |   |------------+------------|                                                                                      */
    /*   |   |  N  | PctN |  N  | PctN |                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |Age|     |      |     |      |                                                                                      */
    /*   |---|     |      |     |      |                                                                                      */
    /*   |11 | 1.00| 11.11| 1.00| 10.00|                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |12 | 2.00| 22.22| 3.00| 30.00|                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |13 | 2.00| 22.22| 1.00| 10.00|                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |14 | 2.00| 22.22| 2.00| 20.00|                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |15 | 2.00| 22.22| 2.00| 20.00|                                                                                      */
    /*   |---+-----+------+-----+------|                                                                                      */
    /*   |16 |    .|     .| 1.00| 10.00|                                                                                      */
    /*   -------------------------------                                                                                      */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
