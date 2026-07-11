************************************************************************
*  Report zidtr_str2json
************************************************************************
*
*  作者：    Mark.WuRY
*  完成日期：
*  描述：    集成组件-结构体转换为报文
************************************************************************
*  版本号 日期   作者   修改描述 功能更改说明书
************************************************************************
*  1.0  2023/01/01  Mark.WuRY   程序创建
************************************************************************
REPORT zidtr_str2json.

TABLES zidt_t_0002.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-h01.
PARAMETERS: p_syid TYPE zidt_t_0002-system_id OBLIGATORY MEMORY ID sir.
PARAMETERS: p_itid TYPE zidt_t_0002-integrate_id OBLIGATORY MEMORY ID cac.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-h02.
  PARAMETERS:
    p_r11 RADIOBUTTON GROUP g1 DEFAULT 'X',
    p_r12 RADIOBUTTON GROUP g1.
SELECTION-SCREEN END OF BLOCK b2.

START-OF-SELECTION.
  PERFORM frm_main.

*&---------------------------------------------------------------------*
*& Form FRM_MAIN
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_main .
  DATA:
    lv_fieldname TYPE string,
    lcl_struc    TYPE REF TO cl_abap_structdescr,
    lcr_struc    TYPE REF TO data,
    lcl_table    TYPE REF TO cl_abap_datadescr,
    lcr_table    TYPE REF TO data.
  FIELD-SYMBOLS:
     <lft_table> TYPE table.

  SELECT SINGLE *
    FROM zidt_t_0002
    INTO @DATA(ls_zidt_t_0002)
   WHERE system_id    = @p_syid
  AND integrate_id = @p_itid.
  IF sy-subrc <> 0.
    MESSAGE s208(00) WITH '数据取得失败' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  CASE abap_true.
    WHEN p_r11.
      SELECT
        tabname,
        fieldname,
        position,
        rollname,
        datatype,
        depth
        INTO TABLE @DATA(lt_dd03l)
        FROM dd03l
       WHERE tabname = @ls_zidt_t_0002-import_structure.
      IF sy-subrc = 0.
        lcl_struc ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_zidt_t_0002-import_structure ).
      ELSE.
        MESSAGE s208(00) WITH '传入结构有误，请检查集成维护视图' DISPLAY LIKE 'E'.
        LEAVE LIST-PROCESSING.
      ENDIF.

    WHEN p_r12.
      SELECT
        tabname,
        fieldname,
        position,
        rollname,
        datatype,
        depth
        INTO TABLE @lt_dd03l
        FROM dd03l
       WHERE tabname = @ls_zidt_t_0002-export_structure.

      IF sy-subrc = 0.
        lcl_struc ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_zidt_t_0002-export_structure ).
      ELSE.
        MESSAGE s208(00) WITH '传出结构有误，请检查集成维护视图护' DISPLAY LIKE 'E'.
        LEAVE LIST-PROCESSING.
      ENDIF.

    WHEN OTHERS.
  ENDCASE.

  CREATE DATA lcr_struc TYPE HANDLE lcl_struc.
  ASSIGN lcr_struc->* TO FIELD-SYMBOL(<lfs_structure>).

  ASSIGN COMPONENT 'IMPORT_PARAMENT-IMPORT_DETAIL-SYSTEM_ID' OF STRUCTURE <lfs_structure> TO FIELD-SYMBOL(<lfs_system_id>).
  IF sy-subrc = 0.
    <lfs_system_id> = p_syid.
  ENDIF.

  ASSIGN COMPONENT 'IMPORT_PARAMENT-IMPORT_DETAIL-INTEGRATE_ID' OF STRUCTURE <lfs_structure> TO FIELD-SYMBOL(<lfs_integrate_id>).
  IF sy-subrc = 0.
    <lfs_integrate_id> = p_itid.
  ENDIF.

  ASSIGN COMPONENT 'IMPORT_PARAMENT-IMPORT_DETAIL-ACCESS_SIGN' OF STRUCTURE <lfs_structure> TO FIELD-SYMBOL(<lfs_access_sign>).
  IF sy-subrc = 0.
    <lfs_access_sign> = ls_zidt_t_0002-access_sign.
  ENDIF.

  LOOP AT lt_dd03l INTO DATA(ls_dd03l) WHERE datatype = 'TTYP'.
    CLEAR lv_fieldname.
    CREATE DATA lcr_table TYPE (ls_dd03l-rollname).
    ASSIGN lcr_table->* TO <lft_table>.

    CREATE DATA lcr_struc LIKE LINE OF <lft_table>.
    ASSIGN lcr_struc->* TO FIELD-SYMBOL(<lfs_table>).

    PERFORM frm_get_structure
      USING ls_dd03l-rollname
   CHANGING <lfs_table>.

    APPEND <lfs_table> TO <lft_table>.

    IF ls_dd03l-depth IS NOT INITIAL.
      DATA(lv_position) = ls_dd03l-position.
      DATA(lv_depth) = ls_dd03l-depth.

      DO ls_dd03l-depth TIMES.
        lv_depth = lv_depth - 1.
        DO.
          lv_position = lv_position - 1.
          IF lv_position = 0.
            EXIT.
          ENDIF.

          READ TABLE lt_dd03l INTO DATA(ls_dd03l_tmp)
            WITH KEY depth    = lv_depth
                     position = lv_position
                     datatype = 'STRU'.
          IF sy-subrc = 0.
            IF lv_fieldname IS INITIAL.
              lv_fieldname = ls_dd03l_tmp-fieldname.
            ELSE.
              lv_fieldname = lv_fieldname && '-' && ls_dd03l_tmp-fieldname.
            ENDIF.
            EXIT.
          ENDIF.
        ENDDO.
      ENDDO.
    ENDIF.

    IF lv_fieldname IS INITIAL.
      lv_fieldname = ls_dd03l-fieldname.
    ELSE.
      lv_fieldname = lv_fieldname && '-' && ls_dd03l-fieldname.
    ENDIF.

    ASSIGN COMPONENT lv_fieldname OF STRUCTURE <lfs_structure> TO FIELD-SYMBOL(<lft_str_table>).
    IF <lft_str_table> IS ASSIGNED.
      <lft_str_table> = <lft_table>.
    ENDIF.

  ENDLOOP.

* 转换结构到JSON
  DATA(lv_json) = /ui2/cl_json=>serialize( data        = <lfs_structure>
                                           pretty_name = /ui2/cl_json=>pretty_mode-none ).

  TRY.
*     将JSON转换为HTML
      CALL TRANSFORMATION sjson2html SOURCE XML lv_json
                                     RESULT XML DATA(lv_html).
    CATCH cx_xslt_runtime_error INTO DATA(lo_err).
      DATA(lv_err_text) = lo_err->get_text( ).
  ENDTRY.

* 显示HTML
  DATA(lv_convert) = cl_abap_codepage=>convert_from( lv_html ).
  cl_abap_browser=>show_html( html_string = lv_convert ).
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_get_structure
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*&      --> UV_STRUCTURE
*&      <-- <LFS_TABLE>
*&---------------------------------------------------------------------*
FORM frm_get_structure
  USING    uv_structure
  CHANGING cs_output.
  DATA:
    lcl_struc TYPE REF TO cl_abap_structdescr,
    lcr_struc TYPE REF TO data,
    lcl_table TYPE REF TO cl_abap_datadescr,
    lcr_table TYPE REF TO data.

  FIELD-SYMBOLS:
    <lft_table> TYPE table.

  SELECT SINGLE
    typename,
    rowtype,
    datatype
    INTO @DATA(ls_dd40l)
    FROM dd40l
   WHERE typename = @uv_structure
     AND as4local = 'A'.
  IF sy-subrc = 0.
    SELECT
      tabname,
      fieldname,
      rollname,
      datatype
      INTO TABLE @DATA(lt_dd03l)
      FROM dd03l
     WHERE tabname  = @ls_dd40l-rowtype
       AND datatype = 'TTYP'.
    IF sy-subrc = 0.

      LOOP AT lt_dd03l INTO DATA(ls_dd03l).

        CREATE DATA lcr_table TYPE (ls_dd03l-rollname).
        ASSIGN lcr_table->* TO <lft_table>.

        CREATE DATA lcr_struc LIKE LINE OF <lft_table>.
        ASSIGN lcr_struc->* TO FIELD-SYMBOL(<lfs_table>).

        PERFORM frm_get_structure
          USING ls_dd03l-rollname
       CHANGING <lfs_table>.

        APPEND <lfs_table> TO <lft_table>.

        ASSIGN COMPONENT ls_dd03l-fieldname OF STRUCTURE cs_output TO FIELD-SYMBOL(<lft_str_table>).
        IF <lft_str_table> IS ASSIGNED.
          <lft_str_table> = <lft_table>.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDIF.
ENDFORM.