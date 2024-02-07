class ZSHAI_EXCEL definition
  public
  final
  create public .

public section.

  methods GENERATE_EXCEL
    importing
      !IV_INFILE type STRING optional .
protected section.
private section.

  class-data GRA_EXCEL type OLE2_OBJECT .
  class-data GRA_SHEETS type OLE2_OBJECT .
  class-data GRA_SHEET type OLE2_OBJECT .
  class-data GRA_CELL type OLE2_OBJECT .
  class-data GRA_FONT type OLE2_OBJECT .
  class-data GRA_WORKSHEET type OLE2_OBJECT .
  class-data GRA_RANGE type OLE2_OBJECT .
  class-data GRA_BORDERS type OLE2_OBJECT .
  class-data GRA_COLUMNS type OLE2_OBJECT .
  class-data GRA_LOADSTATS type OLE2_OBJECT .
  class-data GRA_GEN_ERR type OLE2_OBJECT .
  class-data GV_RC type I .
  class-data GV_S type CHAR1 value CL_ABAP_CHAR_UTILITIES=>HORIZONTAL_TAB ##NO_TEXT.
  class-data GRA_CELL1 type OLE2_OBJECT .
  class-data GRA_CELL2 type OLE2_OBJECT .

  methods CREATE_NEW_SHEET
    importing
      !IV_SHEET type OLE2_OBJECT
      !IV_NAME type STRING
      !IT_DATA type ZRAWFILE_T .
ENDCLASS.



CLASS ZSHAI_EXCEL IMPLEMENTATION.
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
ENDCLASS.
