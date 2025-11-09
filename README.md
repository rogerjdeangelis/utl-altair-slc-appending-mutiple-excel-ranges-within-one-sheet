# utl-altair-slc-appending-mutiple-excel-ranges-within-one-sheet
Altair slc appending mutiple excel ranges within one sheet
    %let pgm=utl-altair-slc-appending-mutiple-excel-ranges-within-one-sheet;

    %stop_submission;

    Altair slc appending mutiple excel ranges within one sheet

    Too long to pst on listserv, see github
    https://github.com/rogerjdeangelis/utl-altair-slc-appending-mutiple-excel-ranges-within-one-sheet

    /community.altair.com
    https://community.altair.com/discussion/5965?tab=accepted

    OPS PROBLEM

    I need to capture the data for each line that represents a loan.  Lines 14-19 and Lines 25-31
    I will need to append the Account number to each line for the account that is showing above each section.
    Lines 14-19 (Append Account 123456) and Lines 25-31 (Append Account 1237890)
    Spreadsheet has extra lines and blank lines in the spreadsheet and is a semi structured spreadsheet in my opinion


    TREE SOLUTIONS
      1  Manually create named ranges (could do this programtically)
          ranges
             account1
             balance1
             account2
             balance2

      2  Hardcode ranges
      3  General solution 100 balances

    REPEATED CUSTOMER BALANCES

    /*****************************************************************************************************************************/
    /* INPUT                                                                                                                     */
    /*                                                                                                                           */
    /* DM0026.11.00            GENERAL LEDGER ACCOUNT BALANCE DETAILS                                                            */
    /*                                                                                                                           */
    /* Database Name: FHBFHP                                                                                                     */
    /*                                                                                                                           */
    /* Report Run Date: 2021-Mar-22 (6:36:57 PM)                                                                                 */
    /*                                                                                                                           */
    /*                                                                                                                           */
    /* Last Process Date: 2021-Mar-22 (12:00:00 AM)                                                                              */
    /*                                                                                                                           */
    /* POSTING UNIT:   0000 - DEFAULT GL UNIT FOR GU AND IN                                                                      */
    /*                                                                                                                           */
    /* ACCOUNT NUMBER:   1234567   DO NOT USEACCT-BADJ-ACBS                                                                      */
    /*                                                                                                                           */
    /*                                                                             PREVIOUS                                      */
    /* CUSTOMER NAME          FAC / LOAN     LIMIT LEVEL  FEE INVESTOR TYPE & NAME BALANCE    DEBITS CREDITS  CLOSING BALANCE    */
    /* ARMSTRONG TRANSFER & L - 92xx4917x                     600 - FIRST HORIZON  47285.77     0      0      47285.77           */
    /* BOB HILSON & COMPANY F - 395xx0321x  00-00-84943543 01 600 - FIRST HORIZON  -249.14      0      0      -249.14            */
    /* HSUS INVESTMENTS, LL L - 92xx4677x                     600 - FIRST HORIZON  1041342.71   0      0      1041342.71         */
    /* NEPHROLOGY ASSOCIATE L - 92xx4458x                     600 - FIRST HORIZON  150037.5     0      0      150037.5           */
    /* ROANE TRANSPORTATION L - 92x5255x                      600 - FIRST HORIZON  1068024.07   0      0      1068024.07         */
    /* U.S. TENNIS AND RECR L - 92xx4862x                     600 - FIRST HORIZON  36371.55     0      0      36371.55           */
    /* ....                                                                                                                      */
    /*                                                                                                                           */
    /*---------------------------------------------------------------------------------------------------------------------------*/
    /*                                                                                                                           */
    /* OUTPUT (APPENDS ALL 100 BALANCE SHEETS)                                                                                   */
    /*                                                                                                                           */
    /*                                                                                                                           */
    /* DATABASE ACCOUNT                                                                        PREVIOUS                CLOSING   */
    /*  NAME    NUMBER    CUSTOMER_NAME      FAC_LOAN     LIMIT_LEVEL  FEE INVESTOR_TYPE_NAME  BALANCE  DEBITS CREDITS BALANCE   */
    /*                                                                                                                           */
    /* FHBFHP 1234567 ARMSTRONG TRANSFER & L-92xx4917x                     600-FIRST HORIZON 47285.77   0       0     47285.77   */
    /* FHBFHP 1234567 BOB HILSON & COMPANY F-395xx0321x 00-00-84943543 01  600-FIRST HORIZON -249.14    0       0     -249.14    */
    /* FHBFHP 1234567 HSUS INVESTMENTS, LL L-92xx4677x                     600-FIRST HORIZON 1041342.71 0       0     1041342.71 */
    /* FHBFHP 1234567 NEPHROLOGY ASSOCIATE L-92xx4458x                     600-FIRST HORIZON 150037.5   0       0     150037.5   */
    /* FHBFHP 1234567 ROANE TRANSPORTATION L-92x5255x                      600-FIRST HORIZON 1068024.07 0       0     1068024.07 */
    /* FHBFHP 1234567 U.S. TENNIS AND RECR L-92xx4862x                     600-FIRST HORIZON 36371.55   0       0     36371.55   */
    /*                                                                                                                           */
    /* FHBFHP 1237890 ARMSTRONG TRANSFER & L-92xx4917x                     600-FIRST HORIZON -47271     0       0     -47271     */
    /* FHBFHP 1237890 BOB HILSON & COMPANY F-395xx0321x 00-00-84943543 01  600-FIRST HORIZON 250        0       0     250        */
    /* FHBFHP 1237890 HSUS INVESTMENTS, LL L-92xx4677x                     600-FIRST HORIZON -1041250   0       0     -1041250   */
    /* FHBFHP 1237890 MAXIMUM TRUCKING LLC F-395xx0200x 00-00-84942701 01  600-FIRST HORIZON 500        0       0     500        */
    /* FHBFHP 1237890 NEPHROLOGY ASSOCIATE L-92xx4458x                     600-FIRST HORIZON -150000    0       0     -150000    */
    /* FHBFHP 1237890 ROANE TRANSPORTATION L-92xx5255x                     600-FIRST HORIZON -1067600   0       0     -1067600   */
    /* FHBFHP 1237890 U.S. TENNIS AND RECR L-92xx4862x                     600-FIRST HORIZON -36362     0       0     -36362     */
    /*                                                                                                                           */
    /*****************************************************************************************************************************/

    DM0026.11.00            GENERAL LEDGER ACCOUNT BALANCE DETAILS

    Database Name: FHBFHP

    Report Run Date: 2021-Mar-22 (6:36:57 PM)


    Last Process Date: 2021-Mar-22 (12:00:00 AM)

    POSTING UNIT:   0000 - DEFAULT GL UNIT FOR GU AND IN

    ACCOUNT NUMBER:   1234567   DO NOT USEACCT-BADJ-ACBS

                                                                                PREVIOUS
    CUSTOMER NAME          FAC / LOAN     LIMIT LEVEL  FEE INVESTOR TYPE & NAME BALANCE    DEBITS CREDITS  CLOSING BALANCE
    ARMSTRONG TRANSFER & L - 92xx4917x                     600 - FIRST HORIZON  47285.77     0      0      47285.77
    BOB HILSON & COMPANY F - 395xx0321x  00-00-84943543 01 600 - FIRST HORIZON  -249.14      0      0      -249.14
    HSUS INVESTMENTS, LL L - 92xx4677x                     600 - FIRST HORIZON  1041342.71   0      0      1041342.71
    NEPHROLOGY ASSOCIATE L - 92xx4458x                     600 - FIRST HORIZON  150037.5     0      0      150037.5
    ROANE TRANSPORTATION L - 92x5255x                      600 - FIRST HORIZON  1068024.07   0      0      1068024.07
    U.S. TENNIS AND RECR L - 92xx4862x                     600 - FIRST HORIZON  36371.55     0      0      36371.55
    ....

    OUTPUT (APPENDS ALL 100 BALANCE SHEETS)


     DATABASE ACCOUNT                                                                           PREVIOUS                 CLOSING_
      NAME    NUMBER    CUSTOMER_NAME       FAC_LOAN       LIMIT_LEVEL FEE INVESTOR_TYPE_NAME  BALANCE   DEBITS CREDITS BALANCE

     FHBFHP  1234567 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON 47285.77    0       0     47285.77
     FHBFHP  1234567 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON -249.14     0       0     -249.14
     FHBFHP  1234567 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON 1041342.71  0       0     1041342.71
     FHBFHP  1234567 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON 150037.5    0       0     150037.5
     FHBFHP  1234567 ROANE TRANSPORTATION  L-92x5255x                       600-FIRST HORIZON 1068024.07  0       0     1068024.07
     FHBFHP  1234567 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON 36371.55    0       0     36371.55

     FHBFHP  1237890 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON -47271      0       0     -47271
     FHBFHP  1237890 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON 250         0       0     250
     FHBFHP  1237890 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON -1041250    0       0     -1041250
     FHBFHP  1237890 MAXIMUM TRUCKING LLC  F-395xx0200x 00-00-84942701  01  600-FIRST HORIZON 500         0       0     500
     FHBFHP  1237890 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON -150000     0       0     -150000
     FHBFHP  1237890 ROANE TRANSPORTATION  L-92xx5255x                      600-FIRST HORIZON -1067600    0       0     -1067600
     FHBFHP  1237890 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON -36362      0       0     -36362


    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */


    libname xls excel "d:/xls/exceltables.xlsx";

    proc print data=xls."sheet1$"n;
    run;quit;

    /***********************************************************************************************************************/
    /* DM0026.11.00            GENERAL LEDGER ACCOUNT BALANCE DETAILS                                                      */
    /*                                                                                                                     */
    /* Database Name: FHBFHP                                                                                               */
    /*                                                                                                                     */
    /* Report Run Date: 2021-Mar-22 (6:36:57 PM)                                                                           */
    /*                                                                                                                     */
    /*                                                                                                                     */
    /* Last Process Date: 2021-Mar-22 (12:00:00 AM)                                                                        */
    /*                                                                                                                     */
    /* POSTING UNIT:   0000 - DEFAULT GL UNIT FOR GU AND IN                                                                */
    /*                                                                                                                     */
    /* ACCOUNT NUMBER:   1234567   DO NOT USEACCT-BADJ-ACBS                                                                */
    /*                                                                                                                     */
    /*                                                                           PREVIOUS                                  */
    /* CUSTOMER NAME        FAC / LOAN    LIMIT LEVEL  FEE INVESTOR TYPE & NAME BALANCE    DEBITS CREDITS  CLOSING BALANCE */
    /* ARMSTRONG TRANSFER & L-92xx4917x                    600 - FIRST HORIZON  47285.77     0      0      47285.77        */
    /* BOB HILSON & COMPANY F-395xx0321x 00-00-84943543 01 600 - FIRST HORIZON  -249.14      0      0      -249.14         */
    /* HSUS INVESTMENTS, LL L-92xx4677x                    600 - FIRST HORIZON  1041342.71   0      0      1041342.71      */
    /* NEPHROLOGY ASSOCIATE L-92xx4458x                    600 - FIRST HORIZON  150037.5     0      0      150037.5        */
    /* ROANE TRANSPORTATION L-92x5255x                     600 - FIRST HORIZON  1068024.07   0      0      1068024.07      */
    /* U.S. TENNIS AND RECR L-92xx4862x                    600 - FIRST HORIZON  36371.55     0      0      36371.55        */
    /* ....                                                                                                                */
    /***********************************************************************************************************************/

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       14:09 Sunday, November  9, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.015
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1         libname xls excel "d:/xls/exceltables.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/exceltables.xlsx

    2
    3         proc print data=xls."sheet1$"n;
    4         run;quit;
    NOTE: 32 observations were read from "XLS.sheet1$"
    NOTE: Procedure print step took :
          real time : 0.345
          cpu time  : 0.296


    5
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.085
          cpu time  : 0.921

    /*                       _                                        _
    / |   ___ _ __ ___  __ _| |_ ___   _ __   __ _ _ __ ___   ___  __| | _ __ __ _ _ __   __ _  ___  ___
    | |  / __| `__/ _ \/ _` | __/ _ \ | `_ \ / _` | `_ ` _ \ / _ \/ _` || `__/ _` | `_ \ / _` |/ _ \/ __|
    | | | (__| | |  __/ (_| | ||  __/ | | | | (_| | | | | | |  __/ (_| || | | (_| | | | | (_| |  __/\__ \
    |_|  \___|_|  \___|\__,_|\__\___| |_| |_|\__,_|_| |_| |_|\___|\__,_||_|  \__,_|_| |_|\__, |\___||___/
                                                                                         |___/
     Manually create named ranges (could do this programtically)
      ranges
         account1
         balance1
         account2
         balance2

    */

    libname xls excel "d:/xls/exceltables.xlsx";

    data want;
     retain f1;
      set xls.account1;
      do until (dne);
      set xls.balance1 end=dne;
      output;
      end;
      set xls.account2;
      do until (dne);
      set xls.balance2 end=dne;
      output;
      end;
      drop f2 f3 f4;
    run;quit;

    proc print data=want;
    run;quit;

    libname xls clear;


    OUTPUT

     ACCOUNT                                                                           PREVIOUS                 CLOSING_
     NUMBER    CUSTOMER_NAME       FAC_LOAN       LIMIT_LEVEL FEE INVESTOR_TYPE_NAME  BALANCE   DEBITS CREDITS BALANCE

    1234567 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON 47285.77    0       0     47285.77
    1234567 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON -249.14     0       0     -249.14
    1234567 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON 1041342.71  0       0     1041342.71
    1234567 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON 150037.5    0       0     150037.5
    1234567 ROANE TRANSPORTATION  L-92x5255x                       600-FIRST HORIZON 1068024.07  0       0     1068024.07
    1234567 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON 36371.55    0       0     36371.55

    1237890 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON -47271      0       0     -47271
    1237890 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON 250         0       0     250
    1237890 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON -1041250    0       0     -1041250
    1237890 MAXIMUM TRUCKING LLC  F-395xx0200x 00-00-84942701  01  600-FIRST HORIZON 500         0       0     500
    1237890 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON -150000     0       0     -150000
    1237890 ROANE TRANSPORTATION  L-92xx5255x                      600-FIRST HORIZON -1067600    0       0     -1067600
    1237890 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON -36362      0       0     -36362

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       14:11 Sunday, November  9, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.032
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1
    2         libname xls excel "d:/xls/exceltables.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/exceltables.xlsx

    3
    4         data want;
    5          retain f1;
    6           set xls.account1;
    7           do until (dne);
    8           set xls.balance1 end=dne;
    9           output;
    10          end;
    11          set xls.account2;
    12          do until (dne);
    13          set xls.balance2 end=dne;
    14          output;
    15          end;
    16          drop f2 f3 f4;
    17        run;

    NOTE: 1 observations were read from "XLS.account1"
    NOTE: 6 observations were read from "XLS.balance1"
    NOTE: 1 observations were read from "XLS.account2"
    NOTE: 7 observations were read from "XLS.balance2"
    NOTE: Data set "WORK.want" has 13 observation(s) and 10 variable(s)
    NOTE: The data step took :
          real time : 0.615
          cpu time  : 0.484


    17      !     quit;

    2                                          Altair SLC       14:11 Sunday, November  9, 2025

    18
    19        proc print data=want;
    20        run;quit;
    NOTE: 13 observations were read from "WORK.want"
    NOTE: Procedure print step took :
          real time : 0.047
          cpu time  : 0.000


    NOTE: Libref XLS has been deassigned.
    21
    22        libname xls clear;
    23
    24
    25
    26
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.417
          cpu time  : 1.140

    /*___    _                   _               _
    |___ \  | |__   __ _ _ __ __| | ___ ___   __| | ___  _ __ __ _ _ __   __ _  ___  ___
      __) | | `_ \ / _` | `__/ _` |/ __/ _ \ / _` |/ _ \| `__/ _` | `_ \ / _` |/ _ \/ __|
     / __/  | | | | (_| | | | (_| | (_| (_) | (_| |  __/| | | (_| | | | | (_| |  __/\__ \
    |_____| |_| |_|\__,_|_|  \__,_|\___\___/ \__,_|\___||_|  \__,_|_| |_|\__, |\___||___/
                                                                         |___/
    */

    libname xls excel "d:/xls/exceltables.xlsx";

    data want;
     retain f1;
      set xls.'sheet1$B10:C11'n ;
      do until (dne);
      set xls.'sheet1$B13:J19'n end=dne;
      output;
      end;
      set xls.'sheet1$B21:C22'n ;
      do until (dne);
      set xls.'sheet1$B24:J31'n end=dne;
      output;
      end;
    run;quit;

    proc print data=want;
    run;quit;


    OUTPUT

     ACCOUNT                                                                           PREVIOUS                 CLOSING_
     NUMBER    CUSTOMER_NAME       FAC_LOAN       LIMIT_LEVEL FEE INVESTOR_TYPE_NAME  BALANCE   DEBITS CREDITS BALANCE

    1234567 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON 47285.77    0       0     47285.77
    1234567 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON -249.14     0       0     -249.14
    1234567 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON 1041342.71  0       0     1041342.71
    1234567 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON 150037.5    0       0     150037.5
    1234567 ROANE TRANSPORTATION  L-92x5255x                       600-FIRST HORIZON 1068024.07  0       0     1068024.07
    1234567 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON 36371.55    0       0     36371.55

    1237890 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON -47271      0       0     -47271
    1237890 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON 250         0       0     250
    1237890 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON -1041250    0       0     -1041250
    1237890 MAXIMUM TRUCKING LLC  F-395xx0200x 00-00-84942701  01  600-FIRST HORIZON 500         0       0     500
    1237890 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON -150000     0       0     -150000
    1237890 ROANE TRANSPORTATION  L-92xx5255x                      600-FIRST HORIZON -1067600    0       0     -1067600
    1237890 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON -36362      0       0     -36362

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       14:14 Sunday, November  9, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.031
          cpu time  : 0.000


    NOTE: AUTOEXEC processing completed

    1
    2         libname xls excel "d:/xls/exceltables.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/exceltables.xlsx

    3
    4         data want;
    5          retain f1;
    6           set xls.'sheet1$B10:C11'n ;
    7           do until (dne);
    8           set xls.'sheet1$B13:J19'n end=dne;
    9           output;
    10          end;
    11          set xls.'sheet1$B21:C22'n ;
    12          do until (dne);
    13          set xls.'sheet1$B24:J31'n end=dne;
    14          output;
    15          end;
    16        run;

    NOTE: 1 observations were read from "XLS.sheet1$B10:C11"
    NOTE: 6 observations were read from "XLS.sheet1$B13:J19"
    NOTE: 1 observations were read from "XLS.sheet1$B21:C22"
    NOTE: 7 observations were read from "XLS.sheet1$B24:J31"
    NOTE: Data set "WORK.want" has 13 observation(s) and 11 variable(s)
    NOTE: The data step took :
          real time : 0.618
          cpu time  : 0.500


    16      !     quit;
    17

    2                                          Altair SLC       14:14 Sunday, November  9, 2025

    18        proc print data=want;
    19        run;quit;
    NOTE: 13 observations were read from "WORK.want"
    NOTE: Procedure print step took :
          real time : 0.049
          cpu time  : 0.000


    20
    21
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.333
          cpu time  : 1.015

    /*____                                               _       _   _
    |___ /    __ _  ___ _ __   ___ _ __ __ _   ___  ___ | |_   _| |_(_) ___  _ __
      |_ \   / _` |/ _ \ `_ \ / _ \ `__/ _` | / __|/ _ \| | | | | __| |/ _ \| `_ \
     ___) | | (_| |  __/ | | |  __/ | | (_| | \__ \ (_) | | |_| | |_| | (_) | | | |
    |____/   \__, |\___|_| |_|\___|_|  \__,_| |___/\___/|_|\__,_|\__|_|\___/|_| |_|
             |___/
    */

    proc import
        out=general
        datafile="d:/xls/exceltables.xlsx"
        dbms=xlsx
        replace;
        getnames=no;
        range="Sheet1$"n;
    run;

    data fix;
      retain database_name account_number;
      set general (rename=(
         VAR1  =  DATABASE
         VAR2  =  CUSTOMER_NAME
         VAR3  =  FAC_LOAN
         VAR4  =  LIMIT_LEVEL
         VAR5  =  FEE
         VAR6  =  INVESTOR_TYPE_NAME
         VAR7  =  PREVIOUS_BALANCE
         VAR8  =  DEBITS
         VAR9  =  CREDITS
         VAR10 =  CLOSING_BALANCE
         ));
      if database      =: 'Database Name:'  then database_name=left(scan(database,2,':'));
      if customer_name =: 'ACCOUNT NUMBER:' then account_number=compress(customer_name,,'kd');
      if not missing (input(debits,?? 12.));
      if not missing(customer_name) then output;
    drop database;
    run;quit;

    proc print data=fix width=min;
    run;quit;

    OUTPUT

     DATABASE ACCOUNT                                                                           PREVIOUS                 CLOSING_
      NAME    NUMBER    CUSTOMER_NAME       FAC_LOAN       LIMIT_LEVEL FEE INVESTOR_TYPE_NAME  BALANCE   DEBITS CREDITS BALANCE

     FHBFHP  1234567 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON 47285.77    0       0     47285.77
     FHBFHP  1234567 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON -249.14     0       0     -249.14
     FHBFHP  1234567 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON 1041342.71  0       0     1041342.71
     FHBFHP  1234567 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON 150037.5    0       0     150037.5
     FHBFHP  1234567 ROANE TRANSPORTATION  L-92x5255x                       600-FIRST HORIZON 1068024.07  0       0     1068024.07
     FHBFHP  1234567 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON 36371.55    0       0     36371.55

     FHBFHP  1237890 ARMSTRONG TRANSFER &  L-92xx4917x                      600-FIRST HORIZON -47271      0       0     -47271
     FHBFHP  1237890 BOB HILSON & COMPANY  F-395xx0321x 00-00-84943543  01  600-FIRST HORIZON 250         0       0     250
     FHBFHP  1237890 HSUS INVESTMENTS, LL  L-92xx4677x                      600-FIRST HORIZON -1041250    0       0     -1041250
     FHBFHP  1237890 MAXIMUM TRUCKING LLC  F-395xx0200x 00-00-84942701  01  600-FIRST HORIZON 500         0       0     500
     FHBFHP  1237890 NEPHROLOGY ASSOCIATE  L-92xx4458x                      600-FIRST HORIZON -150000     0       0     -150000
     FHBFHP  1237890 ROANE TRANSPORTATION  L-92xx5255x                      600-FIRST HORIZON -1067600    0       0     -1067600
     FHBFHP  1237890 U.S. TENNIS AND RECR  L-92xx4862x                      600-FIRST HORIZON -36362      0       0     -36362

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       14:17 Sunday, November  9, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"

    NOTE: 1 record was written to file PRINT

    NOTE: The data step took :
          real time : 0.031
          cpu time  : 0.015


    NOTE: AUTOEXEC processing completed

    1
    2         proc import
    3             out=general
    4             datafile="d:/xls/exceltables.xlsx"
    5             dbms=xlsx
    6             replace;
    7             getnames=no;
    8             range="Sheet1$"n;
    9         run;
    NOTE: Procedure import step took :
          real time : 0.000
          cpu time  : 0.000


    10        libname _XLSXIMP xlsx "d:\xls\exceltables.xlsx" access=readonly
    NOTE: Library _XLSXIMP assigned as follows:
          Engine:        XLSX
          Physical Name: d:\xls\exceltables.xlsx

    11        header=NO
    12        ;
    13        data general;
    14        set _XLSXIMP.'Sheet1$'n;
    15        ;
    16        run;

    NOTE: 32 observations were read from "_XLSXIMP.Sheet1"
    NOTE: Data set "WORK.general" has 32 observation(s) and 10 variable(s)
    NOTE: The data step took :
          real time : 0.000
          cpu time  : 0.015



    2                                          Altair SLC       14:17 Sunday, November  9, 2025

    NOTE: Libref _XLSXIMP has been deassigned.
    17        libname _XLSXIMP clear;
    18
    19        data fix;
    20          retain database_name account_number;
    21          set general (rename=(
    22             VAR1  =  DATABASE
    23             VAR2  =  CUSTOMER_NAME
    24             VAR3  =  FAC_LOAN
    25             VAR4  =  LIMIT_LEVEL
    26             VAR5  =  FEE
    27             VAR6  =  INVESTOR_TYPE_NAME
    28             VAR7  =  PREVIOUS_BALANCE
    29             VAR8  =  DEBITS
    30             VAR9  =  CREDITS
    31             VAR10 =  CLOSING_BALANCE
    32             ));
    33          if database      =: 'Database Name:'  then database_name=left(scan(database,2,':'));
    34          if customer_name =: 'ACCOUNT NUMBER:' then account_number=compress(customer_name,,'k
    34      ! d');
    35          if not missing (input(debits,?? 12.));
    36          if not missing(customer_name) then output;
    37        drop database;
    38        run;

    NOTE: 32 observations were read from "WORK.general"
    NOTE: Data set "WORK.fix" has 13 observation(s) and 11 variable(s)
    NOTE: The data step took :
          real time : 0.015
          cpu time  : 0.000


    38      !     quit;
    39
    40        proc print data=fix width=min;
    41        run;quit;
    NOTE: 13 observations were read from "WORK.fix"
    NOTE: Procedure print step took :
          real time : 0.031
          cpu time  : 0.015


    42
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 0.162
          cpu time  : 0.125

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
