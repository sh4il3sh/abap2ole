CLASS zshai_excel DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS generate_excel
      IMPORTING
        !iv_infile TYPE string OPTIONAL .
    METHODS update_existing_excel .
  PROTECTED SECTION.
  PRIVATE SECTION.

    CLASS-DATA gra_excel TYPE ole2_object .
    CLASS-DATA gra_sheets TYPE ole2_object .
    CLASS-DATA gra_sheet TYPE ole2_object .
    CLASS-DATA gra_cell TYPE ole2_object .
    CLASS-DATA gra_font TYPE ole2_object .
    CLASS-DATA gra_worksheet TYPE ole2_object .
    CLASS-DATA gra_range TYPE ole2_object .
    CLASS-DATA gra_borders TYPE ole2_object .
    CLASS-DATA gra_columns TYPE ole2_object .
    CLASS-DATA gra_loadstats TYPE ole2_object .
    CLASS-DATA gra_gen_err TYPE ole2_object .
    CLASS-DATA gv_rc TYPE i .
    CLASS-DATA gv_s TYPE char1 VALUE cl_abap_char_utilities=>horizontal_tab ##NO_TEXT.
    CLASS-DATA gra_cell1 TYPE ole2_object .
    CLASS-DATA gra_cell2 TYPE ole2_object .

    METHODS create_new_sheet
      IMPORTING
        !iv_sheet TYPE ole2_object
        !iv_name  TYPE string
        !it_data  TYPE zrawfile_t .
ENDCLASS.



CLASS zshai_excel IMPLEMENTATION.
  METHOD create_new_sheet.
    IF it_data IS INITIAL.
      EXIT.
    ENDIF.
    DATA(ira_sheet) = iv_sheet.
    DATA(it_sheet) = it_data.

    IF iv_name <> 'COUNT'.
      " Subsequent sheets to add and paste log
      GET PROPERTY OF gra_excel 'Sheets' = ira_sheet.
      CALL METHOD OF ira_sheet 'Add' = gra_sheet.
      GET PROPERTY OF gra_excel 'ACTIVESHEET' = gra_worksheet.
      SET PROPERTY OF gra_worksheet 'Name' = iv_name. " Sheet name
    ELSE.
      " Create the first sheet with Stats
      GET PROPERTY OF gra_excel 'ACTIVESHEET' = gra_worksheet.
      SET PROPERTY OF gra_worksheet 'Name' = iv_name. " Sheet name
    ENDIF.

    " Copy data in clipboard
    cl_gui_frontend_services=>clipboard_export( IMPORTING  data                 = it_sheet
                                                CHANGING   rc                   = gv_rc
                                                EXCEPTIONS cntl_error           = 1                " Control error
                                                           error_no_gui         = 2                " No GUI available
                                                           not_supported_by_gui = 3                " GUI does not support this
                                                           no_authority         = 4                " Authorization check failed
                                                           OTHERS               = 5 ).
    IF gv_rc <> 0 OR sy-subrc <> 0.
      MESSAGE 'Error Copying data to clipboard'(111) TYPE 'W'.
    ENDIF.

    IF iv_name <> 'COUNT'.
      " Cell Selection 1:1 for subsequent sheets
      CALL METHOD OF gra_excel 'Cells' = gra_cell1
        EXPORTING #1 = 1 " Row
                  #2 = 1. " Column

      CALL METHOD OF gra_excel 'Cells' = gra_cell2
        EXPORTING #1 = 1 " Row
                  #2 = 1. " Column

      CALL METHOD OF gra_excel 'Range' = gra_range
        EXPORTING #1 = gra_cell1
                  #2 = gra_cell2.
    ELSE.
      " Cell Selection custom for first sheet with stats.
      CALL METHOD OF gra_excel 'Cells' = gra_cell1
        EXPORTING #1 = 3 " Row
                  #2 = 3. " Column

      CALL METHOD OF gra_excel 'Cells' = gra_cell2
        EXPORTING #1 = 20 " Row
                  #2 = 9. " Column

      CALL METHOD OF gra_excel 'Range' = gra_range
        EXPORTING #1 = gra_cell1
                  #2 = gra_cell2.
    ENDIF.

    " Above created range is selection in excel file
    CALL METHOD OF gra_range 'Select'.
    " Paste data from clipboard to gra_worksheet.
    CALL METHOD OF gra_worksheet 'Paste'.

    " ---------------------------------------------------------------------
    " Below we do formatting/cosmetic changes for the data pasted------
    " ---------------------------------------------------------------------
    IF iv_name = 'COUNT'.
      " Logic to assign gra_borders to fetched data in worksheet.
      DATA(lv_i) = 3.
      " TODO: variable is assigned but never used (ABAP cleaner)
      LOOP AT it_sheet ASSIGNING FIELD-SYMBOL(<lfs_sheet>).
        DATA(lv_tabix) = sy-tabix.
        lv_i += 1.
        DATA(lv_first)  = |C{ lv_i }|. " Column from where you want to start providing gra_borders.
        DATA(lv_second) = |I{ lv_i }|. " Column up to which you want to provide the gra_borders.

        IF lv_tabix = lines( it_sheet ). " Not add border on last loop
          CONTINUE.
        ENDIF.
        " Make gra_range of selected columns.
        CALL METHOD OF gra_excel 'Range' = gra_range
          EXPORTING #1 = lv_first
                    #2 = lv_second.

        " Logic to assign border on left side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 7. " 7 for left side
        SET PROPERTY OF gra_borders 'LineStyle' = 1. " type of line.

        " Logic to assign border on right side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 8.
        SET PROPERTY OF gra_borders 'LineStyle' = 1.

        " Logic to assign border on top side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 9.
        SET PROPERTY OF gra_borders 'LineStyle' = 1.

        " Logic to assign border on bottom side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 10.
        SET PROPERTY OF gra_borders 'LineStyle' = 1.

        " Logic to assign border on vertical side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 11.
        SET PROPERTY OF gra_borders 'LineStyle' = 1.

        " Logic to assign border on horizontal side.
        CALL METHOD OF gra_range 'Borders' = gra_borders    NO FLUSH
          EXPORTING #1 = 12.
        SET PROPERTY OF gra_borders 'LineStyle' = 1.

        IF lv_tabix = 1.
          GET PROPERTY OF gra_range 'FONT' = gra_font NO FLUSH.
          SET PROPERTY OF gra_font 'BOLD' = 1 NO FLUSH.
          CALL METHOD OF gra_range 'INTERIOR' = gra_range.
          SET PROPERTY OF gra_range 'ColorIndex' = 37.
          SET PROPERTY OF gra_range 'Pattern' = 1.
        ENDIF.

      ENDLOOP.

      CALL METHOD OF gra_worksheet 'Columns' = gra_columns.
      CALL METHOD OF gra_columns 'Autofit'.
    ELSE.
      " For remaining sheets
      DATA(ls_sheet) = it_sheet[ 1 ].
      " We calculate total no of columns based on delimiter
      FIND ALL OCCURRENCES OF gv_s IN ls_sheet RESULTS DATA(lt_result).
      DATA(lv_cols) = lines( lt_result ).

      DO ( lv_cols + 1 ) TIMES.
        lv_i += 1.
        CALL METHOD OF gra_excel 'CELLS' = gra_cell  NO FLUSH
          EXPORTING #1 = 1
                    #2 = lv_i.

        GET PROPERTY OF gra_cell 'FONT' = gra_font NO FLUSH.
        SET PROPERTY OF gra_font 'BOLD' = 1 NO FLUSH.
        CALL METHOD OF gra_cell 'INTERIOR' = gra_range.
        SET PROPERTY OF gra_range 'ColorIndex' = 22.
        SET PROPERTY OF gra_range 'Pattern' = 1.
      ENDDO.
      CALL METHOD OF gra_worksheet 'Columns' = gra_columns.
      CALL METHOD OF gra_columns 'Autofit'.
    ENDIF.
  ENDMETHOD.


  METHOD generate_excel.
    DATA gt_excel TYPE zrawfile_t.

    "Start Excel
    CREATE OBJECT gra_excel 'EXCEL.APPLICATION'.

    "Get list of workbooks, initially empty
    CALL METHOD OF gra_excel 'Workbooks' = gra_sheets.
    SET PROPERTY OF gra_excel 'Visible' = 0.
    CALL METHOD OF gra_sheets 'Add' = gra_sheet.

    create_new_sheet( EXPORTING iv_sheet = gra_loadstats iv_name  = 'COUNT'   it_data  = gt_excel ).
    create_new_sheet( EXPORTING iv_sheet = gra_loadstats iv_name  = 'GENERAL' it_data  = gt_excel ).
    create_new_sheet( EXPORTING iv_sheet = gra_loadstats iv_name  = 'BPID'    it_data  = gt_excel ).

    CALL METHOD OF gra_sheet 'SaveAs'
      EXPORTING
        #1 = iv_infile.
    IF sy-subrc EQ 0.
      MESSAGE 'File downloaded successfully'(110) TYPE 'S'.
    ELSE.
      MESSAGE 'Error downloading the file'(112) TYPE 'W'.
    ENDIF.

    CALL METHOD OF gra_excel 'quit'.

** Free Excel objects
    FREE OBJECT: gra_excel, gra_sheets, gra_sheet, gra_cell, gra_font,
                 gra_worksheet, gra_range, gra_borders.


  ENDMETHOD.

  METHOD update_existing_excel.

    lv_name = 'Replication COUNT'(043).

* START THE EXCEL APPLICATION
    CREATE OBJECT gra_excelz 'EXCEL.APPLICATION'.
    CALL METHOD OF gra_excelz 'Workbooks' = gra_workbooksz.
    SET PROPERTY OF gra_excelz  'Visible' = 0.
    CALL METHOD OF gra_workbooksz 'OPEN'
      EXPORTING
        #1 = p_filez.

    GET PROPERTY OF gra_excelz 'Sheets' = gra_sheetsz.
    CALL METHOD OF gra_sheetsz 'Add' = gra_workbooksz.
    GET PROPERTY OF gra_excelz 'ACTIVESHEET' = gra_sheetsz.
    SET PROPERTY OF gra_sheetsz 'Name' = lv_name. "Sheet name

    CALL METHOD OF gra_excelz 'Worksheets' = gra_worksheetz
     EXPORTING #1 = 1.

    CALL METHOD OF gra_worksheetz 'Activate'.
    IF sy-subrc <> 0.
      MESSAGE 'Error adding Replication sheet'(040) TYPE 'I'.
      CALL METHOD OF gra_excelz 'QUIT'.

      FREE OBJECT: gra_excelz, gra_workbooksz, gra_sheetsz,
                   gra_worksheetz, gra_cell1z, gra_cell2z,  gra_rangez.
      DATA(lv_flag) = 'X'.
    ENDIF.

    IF lv_flag <> 'X'. "This means the existing sheet is accessible.

      "Copy data in clipboard
      cl_gui_frontend_services=>clipboard_export(
        IMPORTING
          data                 = gt_excel
        CHANGING
          rc                   = lv_rc
      EXCEPTIONS
        cntl_error           = 1                " Control error
        error_no_gui         = 2                " No GUI available
        not_supported_by_gui = 3                " GUI does not support this
        no_authority         = 4                " Authorization check failed
        OTHERS               = 5
      ).
      IF ( lv_rc <> 0 OR sy-subrc <> 0 ).
        MESSAGE 'Error Copying data to clipboard'(111) TYPE 'W'.
      ENDIF.

      CALL METHOD OF gra_excelz 'Cells' = gra_cell1z
        EXPORTING
        #1 = 3 "Row
        #2 = 3. "Column

      CALL METHOD OF gra_excelz 'Cells' = gra_cell2z
        EXPORTING
        #1 = 20 "Row
        #2 = 9. "Column

      CALL METHOD OF gra_excelz 'Range' = gra_rangez
        EXPORTING
        #1 = gra_cell1z
        #2 = gra_cell2z.

      CALL METHOD OF gra_rangez 'Select'.
      CALL METHOD OF gra_sheetsz 'Paste'.

      "Begin Formatting sheet
*    DO 6 TIMES.
*      lv_counter += 1.
*      CALL METHOD OF gra_excelz 'CELLS' = gra_cellz  NO FLUSH
*         EXPORTING #1 = 1
*                  #2 = lv_counter .
*
*      GET PROPERTY OF gra_cellz 'FONT' = gra_fontz NO FLUSH.
*      SET PROPERTY OF gra_fontz 'BOLD' = 1 NO FLUSH.
*      CALL METHOD OF gra_cellz 'INTERIOR' = gra_rangez.
*      SET PROPERTY OF gra_rangez 'ColorIndex' = 43.
      "      CALL METHOD OF lo_cell 'Interior' = lo_interior.
      "      SET PROPERTY OF lo_interior 'Color' = 15773696. "Hex color in decimal
*      SET PROPERTY OF gra_rangez 'Pattern' = 1.
*    ENDDO.

      "Logic to assign gra_borders to fetched data in worksheet.
      DATA(lv_i) = 3.
      LOOP AT gt_excel ASSIGNING FIELD-SYMBOL(<lfs_sheet>).
        DATA(lv_tabix) = sy-tabix.
        lv_i = lv_i + 1.
        DATA(lv_first) = |C{ lv_i }|. "Column from where you want to start providing gra_borders.
        DATA(lv_second) = |I{ lv_i }|. "Column up to which you want to provide the gra_borders.

        CHECK lv_tabix <> lines( gt_excel ). "Not add border on last loop

        "Make gra_range of selected columns.
        CALL METHOD OF gra_excelz 'Range' = gra_rangez
          EXPORTING
          #1 = lv_first
          #2 = lv_second.

        "Logic to assign border on left side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
         EXPORTING #1  = 7. "7 for left side
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1. "type of line.

        "Logic to assign border on right side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
           EXPORTING #1  = 8.
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1.

        "Logic to assign border on top side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
           EXPORTING #1  = 9.
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1.

        "Logic to assign border on bottom side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
           EXPORTING #1  = 10.
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1.

        "Logic to assign border on vertical side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
           EXPORTING #1  = 11.
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1.

        "Logic to assign border on horizontal side.
        CALL METHOD OF gra_rangez 'Borders' = gra_bordersz    NO FLUSH
           EXPORTING #1  = 12.
        SET PROPERTY OF gra_bordersz 'LineStyle' = 1.

        IF lv_tabix = 1.
          GET PROPERTY OF gra_rangez 'FONT' = gra_fontz NO FLUSH.
          SET PROPERTY OF gra_fontz 'BOLD' = 1 NO FLUSH.
          CALL METHOD OF gra_rangez 'INTERIOR' = gra_rangez.
          SET PROPERTY OF gra_rangez 'ColorIndex' = 37.
          SET PROPERTY OF gra_rangez 'Pattern' = 1.
        ENDIF.

      ENDLOOP.

      CALL METHOD OF gra_worksheetz 'Columns' = gra_columnsz.
      CALL METHOD OF gra_columnsz 'Autofit'.

      CALL METHOD OF gra_worksheetz 'SAVEAS'
        EXPORTING
          #1 = p_filez
          #2 = 1.
      IF sy-subrc EQ 0.
        MESSAGE 'File replaced successfully'(041) TYPE 'I'.
      ELSE.
        MESSAGE 'Error downloading the file'(042) TYPE 'W'.
      ENDIF.

      CALL METHOD OF gra_excelz 'QUIT'.

      FREE OBJECT: gra_excelz, gra_workbooksz, gra_sheetsz,
                   gra_worksheetz, gra_cell1z, gra_cell2z,  gra_rangez.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
