************************************************************************
*  Report ZIDTR_INTEGRATE_MONITOR
************************************************************************
*
*  作者：    Mark.WuRY
*  完成日期：
*  描述：    集成组件-接口监控
************************************************************************
*  版本号 日期   作者   修改描述 功能更改说明书
************************************************************************
*  1.0  2023/01/01  Mark.WuRY   程序创建
************************************************************************
REPORT zidtr_integrate_monitor.
*&---------------------------------------------------------------------*
*& 包含               全局变量定义
*&---------------------------------------------------------------------*
INCLUDE zidti_abap_top.
INCLUDE zidti_abap_mcr.
INCLUDE zidti_abap_mould.
INCLUDE zidti_abap_form.
INCLUDE zidti_abap_excel.
INCLUDE zidti_abap_bdc.
INCLUDE zidti_abap_event.

TABLES:
  sscrfields.
*&---------------------------------------------------------------------*
* 常量定义:ALV程序中，所有固定值都需要使用声明常量的方式来进行管理
*规范：C_<description>
*&---------------------------------------------------------------------*
*CONSTANTS c_tname            VALUE ''          TYPE tabname.

*&---------------------------------------------------------------------*
* 自定义类型声明:命名规范为TY_<description>
*&---------------------------------------------------------------------*
TYPES:
  ty_field TYPE zidt_t_0003,

  BEGIN OF ty_data,
    integrate_guid TYPE zidt_t_0003-integrate_guid,
    import_datas   TYPE zidt_t_0003-import_datas,
    export_datas   TYPE zidt_t_0003-export_datas,
  END OF   ty_data.

TYPES: BEGIN OF ty_alv.
TYPES:
  box ,
  icon      TYPE  char10,
  mesg      TYPE  string,

  zcolor    TYPE  char4,           "  ALV行颜色
  cell_edit TYPE  lvc_t_styl,      "  单元格编辑
  cellcolor TYPE  lvc_t_scol.      "  单元格颜色

  INCLUDE TYPE ty_field.

TYPES: END OF ty_alv.

*&---------------------------------------------------------------------*
* 全局对象定义：因为全局对象需占用一定程序内存空间，故在定义数据对象时，只
*定义全局内表（GT），全局内表工作区间不需要定义（GS）。如需要可在子程序中声明
*&---------------------------------------------------------------------*
DATA:
  gs_alv   TYPE ty_alv,
  gt_datas TYPE STANDARD TABLE OF ty_data,
  gt_alv   TYPE STANDARD TABLE OF ty_alv,
  go_alv   TYPE REF TO cl_gui_alv_grid.

*&---------------------------------------------------------------------*
*& 包含               选择屏幕
*&---------------------------------------------------------------------*
DATA ls_screen TYPE zidt_t_0003.
*&---------------------------------------------------------------------*
*&  SCREEN
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-h01.
  SELECT-OPTIONS:
    s_guid   FOR ls_screen-integrate_guid,
    s_no     FOR ls_screen-reference_no,
    s_sapno  FOR ls_screen-sap_no,
    s_sid    FOR ls_screen-system_id,
    s_iid    FOR ls_screen-integrate_id,
    s_user   FOR ls_screen-integrate_user,
    s_sign   FOR ls_screen-access_sign,
    s_cdatum FOR ls_screen-cdatum,
    s_cuzeit FOR ls_screen-cuzeit,
    s_cuname FOR ls_screen-cuname.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-h02.
  PARAMETERS:
    r_err RADIOBUTTON GROUP g1,
    r_suc RADIOBUTTON GROUP g1,
    r_all RADIOBUTTON GROUP g1 DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK b2.
*&---------------------------------------------------------------------*
*& 包含               事件类定义及处理
*&---------------------------------------------------------------------*

CLASS zgcl_events_handler DEFINITION.
  PUBLIC SECTION.
    METHODS:
*双击事件方法
      handler_double_click
        FOR EVENT double_click OF cl_gui_alv_grid
        IMPORTING e_row e_column es_row_no,
*数据改变方法
      handler_data_changed
        FOR EVENT data_changed OF cl_gui_alv_grid
        IMPORTING er_data_changed
                  e_onf4
                  e_onf4_before
                  e_onf4_after
                  e_ucomm,
*工具栏设置方法
      handle_toolbar
        FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive,
*按钮事件方法
      handle_user_command
        FOR EVENT user_command OF cl_gui_alv_grid
        IMPORTING e_ucomm,

      onf4
        FOR EVENT onf4 OF cl_gui_alv_grid
        IMPORTING e_fieldname
                  e_fieldvalue
                  es_row_no
                  er_event_data
                  et_bad_cells
                  e_display,

*单机事件方法
      handle_hotspot_click
        FOR EVENT hotspot_click OF cl_gui_alv_grid
        IMPORTING e_row_id e_column_id es_row_no.
ENDCLASS.

CLASS zgcl_events_handler IMPLEMENTATION.
*双击事件方法
  METHOD handler_double_click.
    PERFORM  frm_double_click USING e_row e_column es_row_no.
  ENDMETHOD.                    "HANDLE_NODE_DOUBLE_CLICK
*数据改变方法
  METHOD handler_data_changed .
    PERFORM  frm_data_changed USING  er_data_changed
          e_onf4
          e_onf4_before
          e_onf4_after
          e_ucomm.

  ENDMETHOD.                    "handler_DATA_CHANGED
*工具栏设置方法
  METHOD handle_toolbar.
    PERFORM  frm_set_toolbar CHANGING e_object .
  ENDMETHOD.                    "handle_toolbar
*按钮事件方法
  METHOD handle_user_command.
    PERFORM  frm_user_command USING e_ucomm.
  ENDMETHOD.                    "handle_user_command
*单击事件方法
  METHOD handle_hotspot_click.
    PERFORM  frm_double_click USING e_row_id e_column_id es_row_no.
  ENDMETHOD.                    "handle_hotspot_click

  METHOD onf4 .
    PERFORM  frm_process_event_onf4 USING   e_fieldname
                                       e_fieldvalue
                                       es_row_no
                                       er_event_data
                                       et_bad_cells
                                       e_display.

  ENDMETHOD .                    "HANDLE_ONF4

ENDCLASS.                    "LCL_APPLICATION IMPLEMENTATION

DATA: zgo_handler TYPE REF TO zgcl_events_handler.

FORM  frm_double_click USING pa_row     TYPE  lvc_s_row
                              pa_column  TYPE  lvc_s_col
                              pa_row_no  TYPE  lvc_s_roid.  "#EC NEEDED
  DATA:
    lv_json    TYPE string,
    lv_xml     TYPE string,
    lv_convert TYPE string.
  READ TABLE gt_alv INTO gs_alv INDEX pa_row-index.
  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = gs_alv-integrate_guid.
  CHECK sy-subrc EQ 0.
  SELECT SINGLE data_format INTO @DATA(lv_dataformat) FROM zidt_t_0002
    WHERE system_id    = @gs_alv-system_id
      AND integrate_id = @gs_alv-integrate_id.
  CASE pa_column-fieldname.
    WHEN 'IMPORT_DATAS'.
      CASE lv_dataformat.
        WHEN 'XML'.
          lv_xml  = ls_data-import_datas.
        WHEN OTHERS.
          lv_json = ls_data-import_datas.
      ENDCASE.
    WHEN 'EXPORT_DATAS'.
      IF lv_dataformat = 'XML'.
        lv_xml  = ls_data-export_datas.
      ELSE.
        lv_json = ls_data-export_datas.
      ENDIF.
    WHEN OTHERS.
  ENDCASE.

  IF lv_json IS NOT INITIAL.
    TRY.
*       将JSON转换为HTML
        CALL TRANSFORMATION sjson2html SOURCE XML lv_json
                                       RESULT XML DATA(lv_html).
      CATCH cx_xslt_runtime_error INTO DATA(lo_err).
        DATA(lv_err_text) = lo_err->get_text( ).
    ENDTRY.

    IF lv_html IS INITIAL.

      CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
        EXPORTING
          text   = lv_json
        IMPORTING
          buffer = lv_html
        EXCEPTIONS
          failed = 1
          OTHERS = 2.
      IF sy-subrc <> 0.
* Implement suitable error handling here
      ENDIF.
    ENDIF.

    IF lv_html IS NOT INITIAL.
*   显示HTML
      lv_convert = cl_abap_codepage=>convert_from( lv_html ).
      cl_abap_browser=>show_html( html_string = lv_convert ).
    ENDIF.
  ENDIF.

  IF lv_xml IS NOT INITIAL.
    DATA lv_title TYPE cl_abap_browser=>title.
    lv_title = 'XML报文'.

    cl_abap_browser=>show_xml(
      xml_string = lv_xml
      title      = lv_title
      size       = cl_abap_browser=>medium
    ).

  ENDIF.
ENDFORM.                    " frm_DOUBLE_CLICK

*&---------------------------------------------------------------------*
*&      FORM  FRM_DATA_CHANGED
*&---------------------------------------------------------------------*
*       数据变动
*----------------------------------------------------------------------*
*      -->P_ER_DATA_CHANGED  text
*      -->P_E_ONF4  text
*      -->P_E_UCOMM  text
*----------------------------------------------------------------------*
FORM  frm_data_changed  USING pa_data_changed  TYPE REF TO cl_alv_changed_data_protocol
      pa_onf4	         TYPE	char01
      pa_onf4_before   TYPE char01
      pa_onf4_after	   TYPE	char01
      pa_ucomm         TYPE sy-ucomm.                       "#EC NEEDED

  DATA lv_flg TYPE char1.

*  LOOP AT PA_DATA_CHANGED->MT_GOOD_CELLS INTO DATA(LS_MODI).
*
*    IF LS_MODI-FIELDNAME = 'BOX'.
*      LV_FLG = 'X'.
*      CLEAR GS_ALV.
*      READ TABLE GT_ALV INTO GS_ALV INDEX LS_MODI-ROW_ID.
*      CALL METHOD PA_DATA_CHANGED->GET_CELL_VALUE
*        EXPORTING
*          I_ROW_ID    = LS_MODI-ROW_ID
*          I_FIELDNAME = LS_MODI-FIELDNAME
*        IMPORTING
*          E_VALUE     = GS_ALV-BOX.
*
*      MODIFY GT_ALV FROM GS_ALV
*        TRANSPORTING BOX
*        WHERE MATNR = GS_ALV-MATNR
*          AND WERKS = GS_ALV-WERKS.
*    ENDIF.
*
*  ENDLOOP.
*
*  IF LV_FLG IS NOT INITIAL.
*    "刷新
*    PERFORM FRM_REFRESH_000 USING GO_ALV.
*  ENDIF.

ENDFORM.                    " frm_DATA_CHANGED

FORM  frm_process_event_onf4  USING   e_fieldname TYPE lvc_fname
                                      e_fieldvalue TYPE lvc_value
                                      es_row_no TYPE lvc_s_roid
                                      er_event_data TYPE REF TO cl_alv_event_data
                                      et_bad_cells TYPE lvc_t_modi
                                      e_display TYPE char01. "#EC NEEDED
ENDFORM.                    " PROCESS_EVENT_ONF4

FORM  frm_user_command USING pa_ucomm         TYPE sy-ucomm.
ENDFORM.

*&---------------------------------------------------------------------*
*&      FORM  FRM_SET_TOOLBAR
*&---------------------------------------------------------------------*
*       工具栏设置
*----------------------------------------------------------------------*
*      <--P_E_OBJECT  text
*----------------------------------------------------------------------*
FORM  frm_set_toolbar  CHANGING pa_object TYPE REF TO cl_alv_event_toolbar_set.
*  DATA:ls_button  TYPE stb_button.
*  DATA:lt_button  TYPE ttb_button.
*
*  LOOP AT  pa_object->mt_toolbar INTO  ls_button.
*    APPEND ls_button TO lt_button.
*  ENDLOOP.
*
*
*  pa_object->mt_toolbar = lt_button.
ENDFORM.                    " frm_SET_TOOLBAR
*&---------------------------------------------------------------------*
*& 包含               子历程
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_PROG
*&---------------------------------------------------------------------*
*&      text
*&---------------------------------------------------------------------*
FORM  frm_init_prog .
  PERFORM  frm_init_data.
ENDFORM.                    " FRM_INIT_PROG

*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_DATA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_init_data .
  IF s_cdatum[] IS INITIAL.
    DATA(lv_datum_low) = sy-datum.
    APPEND VALUE #( sign = 'I' option = 'BT' low = lv_datum_low  high = sy-datum  ) TO s_cdatum.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_UPDATE_VARIABLE_VALUE
*&---------------------------------------------------------------------*
*       变量值初始化处理
*----------------------------------------------------------------------*
FORM  frm_update_variable_value .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_SCREENVALUE
*&---------------------------------------------------------------------*
*       屏幕值校验
*----------------------------------------------------------------------*
FORM  frm_check_screenvalue .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SEARCH_HELP
*&---------------------------------------------------------------------*
*       屏幕参数收索帮助
*----------------------------------------------------------------------*
FORM  frm_search_help .

ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  FRM_MODIFY_SCREEN_PARAMETER
*&---------------------------------------------------------------------*
*       屏幕参数调整
*----------------------------------------------------------------------*
FORM  frm_modify_screen_parameter .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_USER_COMMAND_DEAL
*&---------------------------------------------------------------------*
*       用户自定义功能处理
*----------------------------------------------------------------------*
FORM  frm_user_command_deal .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_AUTHORITY_CHECK
*&---------------------------------------------------------------------*
*       权限校验模块
*----------------------------------------------------------------------*
FORM  frm_authority_check .

ENDFORM.
*----------------------------------------------------------------------*
*&      FORM  FRM_PARAMETER_DEAL
*&---------------------------------------------------------------------*
*       数据抽取前参数处理
*----------------------------------------------------------------------*
FORM  frm_parameter_deal .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_GET_DATA
*&---------------------------------------------------------------------*
*       获取相关数据
*----------------------------------------------------------------------*
FORM  frm_data_get .
  DATA ls_data TYPE ty_data.
  CLEAR :
    gt_alv[],
    gt_datas[].

  SELECT *
    INTO TABLE @DATA(lt_0003)
    FROM zidt_t_0003
   WHERE integrate_guid IN @s_guid
     AND reference_no   IN @s_no
     AND sap_no         IN @s_sapno
     AND system_id      IN @s_sid
     AND integrate_id   IN @s_iid
     AND integrate_user IN @s_user
     AND access_sign    IN @s_sign
     AND cdatum         IN @s_cdatum
     AND cuzeit         IN @s_cuzeit
     AND cuname         IN @s_cuname.

  CASE abap_true.
    WHEN r_err.
      DELETE lt_0003 WHERE integrate_type = 'S'.
    WHEN r_suc.
      DELETE lt_0003 WHERE integrate_type <> 'S'.
    WHEN OTHERS.
  ENDCASE.

  LOOP AT lt_0003 INTO DATA(ls_0003).
    CLEAR gs_alv.
    MOVE-CORRESPONDING ls_0003 TO gs_alv.

    IF gs_alv-import_datas IS NOT INITIAL.
      gs_alv-import_datas = '详细'.
    ENDIF.
    IF gs_alv-export_datas IS NOT INITIAL.
      gs_alv-export_datas = '详细'.
    ENDIF.
    APPEND gs_alv TO gt_alv.

    MOVE-CORRESPONDING ls_0003 TO ls_data.
    APPEND ls_data TO gt_datas.
  ENDLOOP.

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_DATA_DEAL
*&---------------------------------------------------------------------*
*       ALV数据处理
*----------------------------------------------------------------------*
FORM  frm_data_deal .

  LOOP AT gt_alv ASSIGNING FIELD-SYMBOL(<fs_alv>).
    IF <fs_alv>-integrate_type = 'S'.
      <fs_alv>-icon = icon_green_light.
    ELSE.
      <fs_alv>-icon = icon_red_light.
    ENDIF.
  ENDLOOP.
  SORT gt_alv BY cdatum DESCENDING cuzeit DESCENDING.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SHOW_ALV
*&---------------------------------------------------------------------*
*       ALV数据展示
*----------------------------------------------------------------------*
FORM  frm_show_alv .

  PERFORM  frm_check_nothing.

  PERFORM  frm_init_fieldcat  CHANGING gt_fieldcat.

  PERFORM  frm_init_comm.

*OO ALV调用
  CALL SCREEN 0100.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_NOTHING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_check_nothing .
  IF gt_alv IS INITIAL .
    PERFORM  frm_get_nothing.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_FIELDCAT
*&---------------------------------------------------------------------*
*       text  初始化ALV输出字段
*----------------------------------------------------------------------*
FORM  frm_init_fieldcat  CHANGING pt_fieldcat  TYPE lvc_t_fcat.
  PERFORM  frm_0100_init_fieldcat CHANGING pt_fieldcat.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_NOTHING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_0100_init_fieldcat  CHANGING pt_fieldcat  TYPE lvc_t_fcat.
  REFRESH:pt_fieldcat.

  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
      i_structure_name       = 'ZIDT_T_0003'
    CHANGING
      ct_fieldcat            = gt_fieldcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.

  mcr_fill_fieldcat:
    'BOX'         '选择',
    'ICON'        '状态',
    'MESG'        '消息'.

  LOOP AT gt_fieldcat ASSIGNING <gfs_fieldcat>.
    <gfs_fieldcat>-tabname = 'GT_ALV'.
    <gfs_fieldcat>-col_pos = sy-tabix.
    <gfs_fieldcat>-col_opt = 'A'.       "设置自动优化时，不生效

    CASE <gfs_fieldcat>-fieldname.
      WHEN 'BOX'.
        <gfs_fieldcat>-checkbox  = c_x.
        <gfs_fieldcat>-edit      = c_x.
        <gfs_fieldcat>-key       = c_x.
        <gfs_fieldcat>-col_opt   = space.
        <gfs_fieldcat>-col_pos   = 1.
        <gfs_fieldcat>-outputlen = '10'.
      WHEN 'ICON'.
        <gfs_fieldcat>-col_pos   = 2.
        <gfs_fieldcat>-key       = c_x.
      WHEN 'MESG'.
        <gfs_fieldcat>-col_pos   = 3.
        <gfs_fieldcat>-key       = c_x.
      WHEN 'INTEGRATE_GUID'.
        <gfs_fieldcat>-col_pos   = <gfs_fieldcat>-col_pos + 3.
        <gfs_fieldcat>-key       = space.
      WHEN 'IMPORT_DATAS' OR 'EXPORT_DATAS'.
        <gfs_fieldcat>-col_pos   = <gfs_fieldcat>-col_pos + 3.
        <gfs_fieldcat>-hotspot   = c_x.
      WHEN OTHERS.
        <gfs_fieldcat>-col_pos = <gfs_fieldcat>-col_pos + 3.
    ENDCASE.
  ENDLOOP.

  pt_fieldcat = gt_fieldcat.

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_COMM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_init_comm .
  gs_layout-zebra      = c_x.
  gs_layout-cwidth_opt = c_x.
  gs_layout-sel_mode   = 'D'.
  gs_layout-no_rowmark = c_x.
  DESCRIBE TABLE gt_alv LINES DATA(lv_lines).
  gv_title = sy-title && '（共' && lv_lines && '条）'.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_STATUS_0100_PBO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_0100_status_pbo .
  DATA:lt_code LIKE TABLE OF sy-ucomm,
       ls_code LIKE sy-ucomm.
  RANGES: lr_ecode FOR sy-ucomm.

  mcr_range_eq: lr_ecode 'ALL'.
  mcr_range_eq: lr_ecode 'SAL'.

  DELETE lr_ecode WHERE low EQ 'ALL'.
  DELETE lr_ecode WHERE low EQ 'SAL'.

  LOOP AT lr_ecode.
    ls_code = lr_ecode-low.
    APPEND ls_code TO lt_code.
  ENDLOOP.

*设置标题与应用工具栏
  SET PF-STATUS 'PF_0100' EXCLUDING lt_code.
  SET TITLEBAR 'TI_0100' WITH gv_title.

*创建并划分ALV容器:前提ALV容易是初始化的
  PERFORM  frm_0100_init.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CREATE_AND_INIT_CONTROLS
*&---------------------------------------------------------------------*
*       ALV容器处理
*----------------------------------------------------------------------*
FORM  frm_0100_init .
  DATA: lv_save,                         "for Parameter I_SAVE-------为变式保存参数
        ls_variant TYPE disvariant.      "for parameter IS_VARIANT---为变式保存参数
  DATA: lt_exclude TYPE ui_functions.    "工具栏排除
  DATA: lv_repid   TYPE sy-repid.        "当前程序号

  IF go_alv IS INITIAL.
*   //变式参数赋值
    lv_save = 'A'.
    PERFORM  frm_get_variant CHANGING ls_variant.
*   //创建ALV控制
    CREATE OBJECT go_alv
      EXPORTING
        i_parent = cl_gui_container=>screen0.
*   //fieldcat设置
    PERFORM  frm_exclude_tb_functions0 CHANGING lt_exclude.
    PERFORM  frm_exclude_tb_functions1 CHANGING lt_exclude.

    gs_layout-info_fname = 'ZCOLOR'.     " 行颜色字段
    gs_layout-ctab_fname = 'CELLCOLOR'.  " 单元格颜色字段
    gs_layout-stylefname = 'CELL_EDIT'.

    CALL METHOD go_alv->set_table_for_first_display
      EXPORTING
        is_variant           = ls_variant
        i_save               = lv_save
        is_layout            = gs_layout
        it_toolbar_excluding = lt_exclude
      CHANGING
        it_outtab            = gt_alv
        it_fieldcatalog      = gt_fieldcat.

    PERFORM  frm_alv_set_handler USING go_alv.
  ELSE.
    PERFORM  frm_refresh_000 USING go_alv.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_GET_VARIANT
*&---------------------------------------------------------------------*
*       ALV通用VARIANT设置
*----------------------------------------------------------------------*
FORM  frm_get_variant  CHANGING pc_variant TYPE disvariant.
  CONCATENATE sy-repid sy-dynnr INTO pc_variant-report.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_ALV_SET_HANDLER
*&---------------------------------------------------------------------*
*       text  OO ALV注册事件
*----------------------------------------------------------------------*
FORM  frm_alv_set_handler USING pa_grid TYPE REF TO cl_gui_alv_grid.
  IF zgo_handler IS INITIAL.
    CREATE OBJECT zgo_handler.
  ENDIF.

  SET HANDLER zgo_handler->handler_double_click FOR pa_grid.
  SET HANDLER zgo_handler->handle_hotspot_click FOR pa_grid.
  SET HANDLER zgo_handler->handle_user_command  FOR pa_grid.
  SET HANDLER zgo_handler->handler_data_changed FOR pa_grid.
  SET HANDLER zgo_handler->handle_toolbar       FOR pa_grid.
  SET HANDLER zgo_handler->onf4                 FOR pa_grid.
  PERFORM  frm_set_grid_onf4 USING   pa_grid.

*  CALL METHOD pa_grid->set_ready_for_input
*    EXPORTING
*      i_ready_for_input = 1.
  CALL METHOD pa_grid->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_enter.      "回车

  CALL METHOD pa_grid->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_modified.  "修改,既光标移出
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SET_GRID_ONF4
*&---------------------------------------------------------------------*
*       text  设置F4
*----------------------------------------------------------------------*
FORM  frm_set_grid_onf4 USING pa_grid TYPE REF TO cl_gui_alv_grid.
  DATA: lv_field TYPE lvc_fname,
        lt_f4    TYPE lvc_t_f4,
        ls_f4    TYPE lvc_s_f4.

  ls_f4-fieldname  = 'MTART'.
  ls_f4-register   = 'X'.
*  ls_f4-getbefore  = 'X'.
*  ls_f4-chngeafter = 'X'.
*  ls_f4-internal   = ''.
  INSERT ls_f4 INTO TABLE lt_f4.

  CALL METHOD pa_grid->register_f4_for_fields
    EXPORTING
      it_f4 = lt_f4.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_USER_COMMAND_0100_PAI
*&---------------------------------------------------------------------*
*       text  用户按键处理
*----------------------------------------------------------------------*
FORM  frm_0100_user_command_pai .
  DATA ls_mse TYPE string.
  gv_code = ok_code.
  CLEAR ok_code.
  CASE gv_code.
    WHEN 'ALL'.   " 全选
      PERFORM  frm_sel_data USING c_x.
    WHEN 'SAL'.
      PERFORM  frm_sel_data USING space.
    WHEN 'REFRESH'.  "
      PERFORM frm_uc_refresh.
    WHEN 'SEND'.  "
      PERFORM frm_uc_send.
  ENDCASE.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SEL_DATA
*&---------------------------------------------------------------------*
*       text  全选/反选
*----------------------------------------------------------------------*
*  -->  pa_mark      text
*----------------------------------------------------------------------*
FORM  frm_sel_data USING pa_mark TYPE c.
  DATA:ls_alv LIKE gs_alv.
  ls_alv-box = pa_mark.
  MODIFY gt_alv FROM ls_alv TRANSPORTING box
  WHERE box NE pa_mark.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_UC_SAVE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_uc_send.
  LOOP AT gt_alv ASSIGNING FIELD-SYMBOL(<lfs_alv>)
    WHERE box  IS NOT INITIAL.
    CLEAR:
      <lfs_alv>-icon,
      <lfs_alv>-mesg.
    IF <lfs_alv>-import_datas IS INITIAL.
      <lfs_alv>-icon = icon_red_light.
      <lfs_alv>-mesg = '无导入报文，不能发送'.
    ENDIF.
  ENDLOOP.

  LOOP AT gt_alv ASSIGNING <lfs_alv>
    WHERE box  IS NOT INITIAL
      AND icon IS INITIAL.
    IF <lfs_alv>-access_sign = 'C'. "Consume Interface
      PERFORM frm_send_c CHANGING <lfs_alv> .
    ENDIF.

    IF <lfs_alv>-access_sign = 'S'."Supplier Interface
      PERFORM frm_send_s CHANGING <lfs_alv>.
    ENDIF.

    IF <lfs_alv>-send_guid IS NOT INITIAL.

      <lfs_alv>-icon = icon_green_light.
      <lfs_alv>-mesg = '发送成功'.

      UPDATE zidt_t_0003
         SET send_guid = <lfs_alv>-integrate_guid
       WHERE integrate_guid = <lfs_alv>-send_guid.
    ENDIF.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_send_c
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_send_c CHANGING cs_alv TYPE ty_alv.
  DATA:
    lv_user   TYPE ZIDT_E_USERNAME,
    lv_url    TYPE string,
    lv_token  TYPE string,
    ls_import TYPE zidt_s_import,
    ls_export TYPE zidt_s_export.

  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = cs_alv-integrate_guid.
  TRY .
      cs_alv-send_guid = cl_uuid_factory=>create_system_uuid( )->create_uuid_x16( ).
    CATCH cx_uuid_error.

  ENDTRY.

  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
   WHERE system_id    = @cs_alv-system_id
     AND integrate_id = @cs_alv-integrate_id.
  IF sy-subrc = 0.
    lv_url = ls_integrate-url.
*   需求TOKEN
    SEARCH lv_url FOR 'token='.
    IF sy-subrc = 0.
*      IF cs_alv-system_id = 'OA'.
*       SAP账户对应OA账户
*        SELECT SINGLE usr21~bname,usr21~persnumber,sort1,zbc_t_0001~loginname INTO @DATA(ls_user_info) FROM usr21
*          INNER JOIN adrp ON adrp~persnumber = usr21~persnumber
*          INNER JOIN zbc_t_0001 ON zbc_t_0001~code = adrp~sort1
*          WHERE usr21~bname = @cs_alv-integrate_user
*            AND zbc_t_0001~state = '1'    "在职
*            AND zbc_t_0001~isdeleted = @space. "删除状态为空
*      ENDIF.
*     无账号映射，使用SAP账户发送
*      IF ls_user_info IS INITIAL.
        lv_user = cs_alv-integrate_user.
*      ELSE.
*        lv_user = ls_user_info-loginname.
*      ENDIF.

      CALL METHOD zidt_c_integrate_tool=>get_token_from_outsyst
        EXPORTING
          iv_system_id       = cs_alv-system_id
          iv_user            = lv_user
        IMPORTING
          ev_token           = lv_token
        EXCEPTIONS
          no_token_configure = 1
          OTHERS             = 2.
      IF sy-subrc <> 0.
*         Implement suitable error handling here
      ENDIF.
      lv_url = lv_url && lv_token.
    ENDIF.
    MOVE-CORRESPONDING ls_integrate TO ls_import-import_parament-import_detail.
    ls_import-import_parament-integrate_guid = cs_alv-send_guid.
    ls_import-import_parament-import_detail-integrate_language = sy-langu.
    ls_import-import_parament-import_detail-integrate_user = sy-uname.
    ls_import-import_datas = ls_data-import_datas.
  ENDIF.

  IF lv_url IS NOT INITIAL.
*   写入导入记录
    CALL METHOD zidt_c_integrate_tool=>write_import_log
      EXPORTING
        is_import_parament = ls_import-import_parament
        iv_import_datas    = ls_import-import_datas.
*   发送HTTP
    CALL FUNCTION 'ZIDT_SET_HTTP'
      EXPORTING
        iv_url    = lv_url
        is_import = ls_import
      IMPORTING
        es_export = ls_export.

    ls_export-export_parament-reference_no = cs_alv-reference_no.
    ls_export-export_parament-sap_no = cs_alv-sap_no.
*   写入导出记录
    CALL METHOD zidt_c_integrate_tool=>write_export_log
      EXPORTING
        is_export_parament = ls_export-export_parament
        iv_export_datas    = ls_export-export_datas.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_send_s
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_send_s  CHANGING cs_alv TYPE ty_alv.
  DATA:
    ls_import  TYPE zidt_s_import,
    ls_export  TYPE zidt_s_export,
    lcl_import TYPE REF TO cl_abap_structdescr,
    lcr_import TYPE REF TO data,
    lcl_export TYPE REF TO cl_abap_structdescr,
    lcr_export TYPE REF TO data.

  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = cs_alv-integrate_guid.

  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
   WHERE system_id    = @cs_alv-system_id
     AND integrate_id = @cs_alv-integrate_id.

  MOVE-CORRESPONDING cs_alv TO ls_import-import_parament-import_detail.
  TRY .
      ls_import-import_parament-integrate_guid = cl_uuid_factory=>create_system_uuid( )->create_uuid_x16( ).
    CATCH cx_uuid_error.

  ENDTRY.
  cs_alv-send_guid = ls_import-import_parament-integrate_guid.
* 写入导入记录
  CALL METHOD zidt_c_integrate_tool=>write_import_log
    EXPORTING
      is_import_parament = ls_import-import_parament
      iv_import_datas    = ls_data-import_datas.
  MOVE-CORRESPONDING ls_import-import_parament TO ls_export-export_parament.
* 调用配置函数
  IF ls_export-export_parament-export_detail-integrate_type IS INITIAL.
    TRY .
        lcl_import ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_integrate-import_structure ).
        CREATE DATA lcr_import TYPE HANDLE lcl_import.
        ASSIGN lcr_import->* TO FIELD-SYMBOL(<lfs_import>).

        lcl_export ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_integrate-export_structure ).
        CREATE DATA lcr_export TYPE HANDLE lcl_export.
        ASSIGN lcr_export->* TO FIELD-SYMBOL(<lfs_export>).
      CATCH  cx_sy_move_cast_error.

    ENDTRY.
    IF <lfs_import> IS ASSIGNED
      AND <lfs_export> IS ASSIGNED.
*       转换JSON到结构话数据
      /ui2/cl_json=>deserialize( EXPORTING json        = ls_data-import_datas
                                           pretty_name = /ui2/cl_json=>pretty_mode-camel_case
                                 CHANGING  data        = <lfs_import> ).
      ASSIGN COMPONENT 'IMPORT_PARAMENT' OF STRUCTURE <lfs_import> TO FIELD-SYMBOL(<lfs_import_parament>).
      MOVE-CORRESPONDING ls_import-import_parament TO <lfs_import_parament>.
      CALL FUNCTION ls_integrate-function
        EXPORTING
          is_import = <lfs_import>
        IMPORTING
          es_export = <lfs_export>.
    ENDIF.
  ENDIF.

  IF <lfs_export> IS ASSIGNED.
    DATA(ev_json) = /ui2/cl_json=>serialize( data        = <lfs_export>
                                             pretty_name = /ui2/cl_json=>pretty_mode-none ).

    ASSIGN COMPONENT 'EXPORT_PARAMENT' OF STRUCTURE <lfs_export> TO FIELD-SYMBOL(<lfs_export_parament>).
    MOVE-CORRESPONDING <lfs_export_parament> TO ls_export-export_parament.
  ELSE.
    ev_json = /ui2/cl_json=>serialize( data        = ls_export
                                       pretty_name = /ui2/cl_json=>pretty_mode-none ).

  ENDIF.
  IF ls_export-export_parament-integrate_guid IS INITIAL.
    ls_export-export_parament-integrate_guid = ls_import-import_parament-integrate_guid.
  ENDIF.
* 写入导出记录
  CALL METHOD zidt_c_integrate_tool=>write_export_log
    EXPORTING
      is_export_parament = ls_export-export_parament
      iv_export_datas    = ev_json.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_get_token
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_get_token  USING    uv_system_id
                             uv_user
                    CHANGING cv_token.
*  DATA:
*    lv_iid       TYPE string,
*    ls_token_in  TYPE zoa_s0000_import,
*    ls_token_out TYPE zstoatokenr,
*    ls_import    TYPE zidt_s_import,
*    ls_export    TYPE zidt_s_export.
*  lv_iid = uv_system_id && '_TOKEN'.
*  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
*   WHERE system_id    = @uv_system_id
*     AND integrate_id = @lv_iid.
** SAP账户对应OA账户
*  SELECT SINGLE usr21~bname,usr21~persnumber,sort1,zbc_t_0001~loginname INTO @DATA(ls_user_info) FROM usr21
*    INNER JOIN adrp ON adrp~persnumber = usr21~persnumber
*    INNER JOIN zbc_t_0001 ON zbc_t_0001~code = adrp~sort1
*    WHERE usr21~bname = @uv_user
*      AND zbc_t_0001~state = '1'    "在职
*      AND zbc_t_0001~isdeleted = @space. "删除状态为空
*
*  ls_token_in-user_name  = ls_integrate-username.
*  ls_token_in-password   = ls_integrate-password.
*  ls_token_in-login_name = ls_user_info-loginname.
*
*  DATA(lv_token_json) = /ui2/cl_json=>serialize( data        = ls_token_in
*                                                 compress    = abap_true
*                                                 pretty_name = /ui2/cl_json=>pretty_mode-camel_case ).
*
*  MOVE-CORRESPONDING ls_integrate TO ls_import-import_parament-import_detail.
*  ls_import-import_datas = lv_token_json.
*
** 发送HTTP
*  CALL FUNCTION 'ZIDT_SET_HTTP'
*    EXPORTING
*      is_import = ls_import
*    IMPORTING
*      es_export = ls_export.
*  IF ls_export-export_datas IS NOT INITIAL.
*    /ui2/cl_json=>deserialize( EXPORTING json        = ls_export-export_datas
*                                         pretty_name = /ui2/cl_json=>pretty_mode-low_case
*                               CHANGING  data        = ls_token_out ).
*  ENDIF.
*  cv_token = ls_token_out-id.
*  IF cv_token IS INITIAL.
*    MESSAGE e208(00) WITH TEXT-m01.
*  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_uc_refresh
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_uc_refresh .

  PERFORM frm_data_get.

  PERFORM frm_data_deal.
  CLEAR gv_title.
  DESCRIBE TABLE gt_alv LINES DATA(lv_lines).
  gv_title = sy-title+0(9) && '（共' && lv_lines && '条）'.
ENDFORM.
************************************************************************
*  Report ZIDTR_INTEGRATE_MONITOR
************************************************************************
*
*  作者：    Mark.WuRY
*  完成日期：
*  描述：    集成组件-接口监控
************************************************************************
*  版本号 日期   作者   修改描述 功能更改说明书
************************************************************************
*  1.0  2023/01/01  Mark.WuRY   程序创建
************************************************************************
REPORT zidtr_integrate_monitor.
*&---------------------------------------------------------------------*
*& 包含               全局变量定义
*&---------------------------------------------------------------------*
INCLUDE zidti_abap_top.
INCLUDE zidti_abap_mcr.
INCLUDE zidti_abap_mould.
INCLUDE zidti_abap_form.
INCLUDE zidti_abap_excel.
INCLUDE zidti_abap_bdc.
INCLUDE zidti_abap_event.

TABLES:
  sscrfields.
*&---------------------------------------------------------------------*
* 常量定义:ALV程序中，所有固定值都需要使用声明常量的方式来进行管理
*规范：C_<description>
*&---------------------------------------------------------------------*
*CONSTANTS c_tname            VALUE ''          TYPE tabname.

*&---------------------------------------------------------------------*
* 自定义类型声明:命名规范为TY_<description>
*&---------------------------------------------------------------------*
TYPES:
  ty_field TYPE zidt_t_0003,

  BEGIN OF ty_data,
    integrate_guid TYPE zidt_t_0003-integrate_guid,
    import_datas   TYPE zidt_t_0003-import_datas,
    export_datas   TYPE zidt_t_0003-export_datas,
  END OF   ty_data.

TYPES: BEGIN OF ty_alv.
TYPES:
  box ,
  icon      TYPE  char10,
  mesg      TYPE  string,

  zcolor    TYPE  char4,           "  ALV行颜色
  cell_edit TYPE  lvc_t_styl,      "  单元格编辑
  cellcolor TYPE  lvc_t_scol.      "  单元格颜色

  INCLUDE TYPE ty_field.

TYPES: END OF ty_alv.

*&---------------------------------------------------------------------*
* 全局对象定义：因为全局对象需占用一定程序内存空间，故在定义数据对象时，只
*定义全局内表（GT），全局内表工作区间不需要定义（GS）。如需要可在子程序中声明
*&---------------------------------------------------------------------*
DATA:
  gs_alv   TYPE ty_alv,
  gt_datas TYPE STANDARD TABLE OF ty_data,
  gt_alv   TYPE STANDARD TABLE OF ty_alv,
  go_alv   TYPE REF TO cl_gui_alv_grid.

*&---------------------------------------------------------------------*
*& 包含               选择屏幕
*&---------------------------------------------------------------------*
DATA ls_screen TYPE zidt_t_0003.
*&---------------------------------------------------------------------*
*&  SCREEN
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-h01.
  SELECT-OPTIONS:
    s_guid   FOR ls_screen-integrate_guid,
    s_no     FOR ls_screen-reference_no,
    s_sapno  FOR ls_screen-sap_no,
    s_sid    FOR ls_screen-system_id,
    s_iid    FOR ls_screen-integrate_id,
    s_user   FOR ls_screen-integrate_user,
    s_sign   FOR ls_screen-access_sign,
    s_cdatum FOR ls_screen-cdatum,
    s_cuzeit FOR ls_screen-cuzeit,
    s_cuname FOR ls_screen-cuname.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-h02.
  PARAMETERS:
    r_err RADIOBUTTON GROUP g1,
    r_suc RADIOBUTTON GROUP g1,
    r_all RADIOBUTTON GROUP g1 DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK b2.
*&---------------------------------------------------------------------*
*& 包含               事件类定义及处理
*&---------------------------------------------------------------------*

CLASS zgcl_events_handler DEFINITION.
  PUBLIC SECTION.
    METHODS:
*双击事件方法
      handler_double_click
        FOR EVENT double_click OF cl_gui_alv_grid
        IMPORTING e_row e_column es_row_no,
*数据改变方法
      handler_data_changed
        FOR EVENT data_changed OF cl_gui_alv_grid
        IMPORTING er_data_changed
                  e_onf4
                  e_onf4_before
                  e_onf4_after
                  e_ucomm,
*工具栏设置方法
      handle_toolbar
        FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive,
*按钮事件方法
      handle_user_command
        FOR EVENT user_command OF cl_gui_alv_grid
        IMPORTING e_ucomm,

      onf4
        FOR EVENT onf4 OF cl_gui_alv_grid
        IMPORTING e_fieldname
                  e_fieldvalue
                  es_row_no
                  er_event_data
                  et_bad_cells
                  e_display,

*单机事件方法
      handle_hotspot_click
        FOR EVENT hotspot_click OF cl_gui_alv_grid
        IMPORTING e_row_id e_column_id es_row_no.
ENDCLASS.

CLASS zgcl_events_handler IMPLEMENTATION.
*双击事件方法
  METHOD handler_double_click.
    PERFORM  frm_double_click USING e_row e_column es_row_no.
  ENDMETHOD.                    "HANDLE_NODE_DOUBLE_CLICK
*数据改变方法
  METHOD handler_data_changed .
    PERFORM  frm_data_changed USING  er_data_changed
          e_onf4
          e_onf4_before
          e_onf4_after
          e_ucomm.

  ENDMETHOD.                    "handler_DATA_CHANGED
*工具栏设置方法
  METHOD handle_toolbar.
    PERFORM  frm_set_toolbar CHANGING e_object .
  ENDMETHOD.                    "handle_toolbar
*按钮事件方法
  METHOD handle_user_command.
    PERFORM  frm_user_command USING e_ucomm.
  ENDMETHOD.                    "handle_user_command
*单击事件方法
  METHOD handle_hotspot_click.
    PERFORM  frm_double_click USING e_row_id e_column_id es_row_no.
  ENDMETHOD.                    "handle_hotspot_click

  METHOD onf4 .
    PERFORM  frm_process_event_onf4 USING   e_fieldname
                                       e_fieldvalue
                                       es_row_no
                                       er_event_data
                                       et_bad_cells
                                       e_display.

  ENDMETHOD .                    "HANDLE_ONF4

ENDCLASS.                    "LCL_APPLICATION IMPLEMENTATION

DATA: zgo_handler TYPE REF TO zgcl_events_handler.

FORM  frm_double_click USING pa_row     TYPE  lvc_s_row
                              pa_column  TYPE  lvc_s_col
                              pa_row_no  TYPE  lvc_s_roid.  "#EC NEEDED
  DATA:
    lv_json    TYPE string,
    lv_xml     TYPE string,
    lv_convert TYPE string.
  READ TABLE gt_alv INTO gs_alv INDEX pa_row-index.
  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = gs_alv-integrate_guid.
  CHECK sy-subrc EQ 0.
  SELECT SINGLE data_format INTO @DATA(lv_dataformat) FROM zidt_t_0002
    WHERE system_id    = @gs_alv-system_id
      AND integrate_id = @gs_alv-integrate_id.
  CASE pa_column-fieldname.
    WHEN 'IMPORT_DATAS'.
      CASE lv_dataformat.
        WHEN 'XML'.
          lv_xml  = ls_data-import_datas.
        WHEN OTHERS.
          lv_json = ls_data-import_datas.
      ENDCASE.
    WHEN 'EXPORT_DATAS'.
      IF lv_dataformat = 'XML'.
        lv_xml  = ls_data-export_datas.
      ELSE.
        lv_json = ls_data-export_datas.
      ENDIF.
    WHEN OTHERS.
  ENDCASE.

  IF lv_json IS NOT INITIAL.
    TRY.
*       将JSON转换为HTML
        CALL TRANSFORMATION sjson2html SOURCE XML lv_json
                                       RESULT XML DATA(lv_html).
      CATCH cx_xslt_runtime_error INTO DATA(lo_err).
        DATA(lv_err_text) = lo_err->get_text( ).
    ENDTRY.

    IF lv_html IS INITIAL.

      CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
        EXPORTING
          text   = lv_json
        IMPORTING
          buffer = lv_html
        EXCEPTIONS
          failed = 1
          OTHERS = 2.
      IF sy-subrc <> 0.
* Implement suitable error handling here
      ENDIF.
    ENDIF.

    IF lv_html IS NOT INITIAL.
*   显示HTML
      lv_convert = cl_abap_codepage=>convert_from( lv_html ).
      cl_abap_browser=>show_html( html_string = lv_convert ).
    ENDIF.
  ENDIF.

  IF lv_xml IS NOT INITIAL.
    DATA lv_title TYPE cl_abap_browser=>title.
    lv_title = 'XML报文'.

    cl_abap_browser=>show_xml(
      xml_string = lv_xml
      title      = lv_title
      size       = cl_abap_browser=>medium
    ).

  ENDIF.
ENDFORM.                    " frm_DOUBLE_CLICK

*&---------------------------------------------------------------------*
*&      FORM  FRM_DATA_CHANGED
*&---------------------------------------------------------------------*
*       数据变动
*----------------------------------------------------------------------*
*      -->P_ER_DATA_CHANGED  text
*      -->P_E_ONF4  text
*      -->P_E_UCOMM  text
*----------------------------------------------------------------------*
FORM  frm_data_changed  USING pa_data_changed  TYPE REF TO cl_alv_changed_data_protocol
      pa_onf4	         TYPE	char01
      pa_onf4_before   TYPE char01
      pa_onf4_after	   TYPE	char01
      pa_ucomm         TYPE sy-ucomm.                       "#EC NEEDED

  DATA lv_flg TYPE char1.

*  LOOP AT PA_DATA_CHANGED->MT_GOOD_CELLS INTO DATA(LS_MODI).
*
*    IF LS_MODI-FIELDNAME = 'BOX'.
*      LV_FLG = 'X'.
*      CLEAR GS_ALV.
*      READ TABLE GT_ALV INTO GS_ALV INDEX LS_MODI-ROW_ID.
*      CALL METHOD PA_DATA_CHANGED->GET_CELL_VALUE
*        EXPORTING
*          I_ROW_ID    = LS_MODI-ROW_ID
*          I_FIELDNAME = LS_MODI-FIELDNAME
*        IMPORTING
*          E_VALUE     = GS_ALV-BOX.
*
*      MODIFY GT_ALV FROM GS_ALV
*        TRANSPORTING BOX
*        WHERE MATNR = GS_ALV-MATNR
*          AND WERKS = GS_ALV-WERKS.
*    ENDIF.
*
*  ENDLOOP.
*
*  IF LV_FLG IS NOT INITIAL.
*    "刷新
*    PERFORM FRM_REFRESH_000 USING GO_ALV.
*  ENDIF.

ENDFORM.                    " frm_DATA_CHANGED

FORM  frm_process_event_onf4  USING   e_fieldname TYPE lvc_fname
                                      e_fieldvalue TYPE lvc_value
                                      es_row_no TYPE lvc_s_roid
                                      er_event_data TYPE REF TO cl_alv_event_data
                                      et_bad_cells TYPE lvc_t_modi
                                      e_display TYPE char01. "#EC NEEDED
ENDFORM.                    " PROCESS_EVENT_ONF4

FORM  frm_user_command USING pa_ucomm         TYPE sy-ucomm.
ENDFORM.

*&---------------------------------------------------------------------*
*&      FORM  FRM_SET_TOOLBAR
*&---------------------------------------------------------------------*
*       工具栏设置
*----------------------------------------------------------------------*
*      <--P_E_OBJECT  text
*----------------------------------------------------------------------*
FORM  frm_set_toolbar  CHANGING pa_object TYPE REF TO cl_alv_event_toolbar_set.
*  DATA:ls_button  TYPE stb_button.
*  DATA:lt_button  TYPE ttb_button.
*
*  LOOP AT  pa_object->mt_toolbar INTO  ls_button.
*    APPEND ls_button TO lt_button.
*  ENDLOOP.
*
*
*  pa_object->mt_toolbar = lt_button.
ENDFORM.                    " frm_SET_TOOLBAR
*&---------------------------------------------------------------------*
*& 包含               子历程
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_PROG
*&---------------------------------------------------------------------*
*&      text
*&---------------------------------------------------------------------*
FORM  frm_init_prog .
  PERFORM  frm_init_data.
ENDFORM.                    " FRM_INIT_PROG

*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_DATA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_init_data .
  IF s_cdatum[] IS INITIAL.
    DATA(lv_datum_low) = sy-datum.
    APPEND VALUE #( sign = 'I' option = 'BT' low = lv_datum_low  high = sy-datum  ) TO s_cdatum.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_UPDATE_VARIABLE_VALUE
*&---------------------------------------------------------------------*
*       变量值初始化处理
*----------------------------------------------------------------------*
FORM  frm_update_variable_value .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_SCREENVALUE
*&---------------------------------------------------------------------*
*       屏幕值校验
*----------------------------------------------------------------------*
FORM  frm_check_screenvalue .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SEARCH_HELP
*&---------------------------------------------------------------------*
*       屏幕参数收索帮助
*----------------------------------------------------------------------*
FORM  frm_search_help .

ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  FRM_MODIFY_SCREEN_PARAMETER
*&---------------------------------------------------------------------*
*       屏幕参数调整
*----------------------------------------------------------------------*
FORM  frm_modify_screen_parameter .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_USER_COMMAND_DEAL
*&---------------------------------------------------------------------*
*       用户自定义功能处理
*----------------------------------------------------------------------*
FORM  frm_user_command_deal .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_AUTHORITY_CHECK
*&---------------------------------------------------------------------*
*       权限校验模块
*----------------------------------------------------------------------*
FORM  frm_authority_check .

ENDFORM.
*----------------------------------------------------------------------*
*&      FORM  FRM_PARAMETER_DEAL
*&---------------------------------------------------------------------*
*       数据抽取前参数处理
*----------------------------------------------------------------------*
FORM  frm_parameter_deal .

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_GET_DATA
*&---------------------------------------------------------------------*
*       获取相关数据
*----------------------------------------------------------------------*
FORM  frm_data_get .
  DATA ls_data TYPE ty_data.
  CLEAR :
    gt_alv[],
    gt_datas[].

  SELECT *
    INTO TABLE @DATA(lt_0003)
    FROM zidt_t_0003
   WHERE integrate_guid IN @s_guid
     AND reference_no   IN @s_no
     AND sap_no         IN @s_sapno
     AND system_id      IN @s_sid
     AND integrate_id   IN @s_iid
     AND integrate_user IN @s_user
     AND access_sign    IN @s_sign
     AND cdatum         IN @s_cdatum
     AND cuzeit         IN @s_cuzeit
     AND cuname         IN @s_cuname.

  CASE abap_true.
    WHEN r_err.
      DELETE lt_0003 WHERE integrate_type = 'S'.
    WHEN r_suc.
      DELETE lt_0003 WHERE integrate_type <> 'S'.
    WHEN OTHERS.
  ENDCASE.

  LOOP AT lt_0003 INTO DATA(ls_0003).
    CLEAR gs_alv.
    MOVE-CORRESPONDING ls_0003 TO gs_alv.

    IF gs_alv-import_datas IS NOT INITIAL.
      gs_alv-import_datas = '详细'.
    ENDIF.
    IF gs_alv-export_datas IS NOT INITIAL.
      gs_alv-export_datas = '详细'.
    ENDIF.
    APPEND gs_alv TO gt_alv.

    MOVE-CORRESPONDING ls_0003 TO ls_data.
    APPEND ls_data TO gt_datas.
  ENDLOOP.

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_DATA_DEAL
*&---------------------------------------------------------------------*
*       ALV数据处理
*----------------------------------------------------------------------*
FORM  frm_data_deal .

  LOOP AT gt_alv ASSIGNING FIELD-SYMBOL(<fs_alv>).
    IF <fs_alv>-integrate_type = 'S'.
      <fs_alv>-icon = icon_green_light.
    ELSE.
      <fs_alv>-icon = icon_red_light.
    ENDIF.
  ENDLOOP.
  SORT gt_alv BY cdatum DESCENDING cuzeit DESCENDING.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SHOW_ALV
*&---------------------------------------------------------------------*
*       ALV数据展示
*----------------------------------------------------------------------*
FORM  frm_show_alv .

  PERFORM  frm_check_nothing.

  PERFORM  frm_init_fieldcat  CHANGING gt_fieldcat.

  PERFORM  frm_init_comm.

*OO ALV调用
  CALL SCREEN 0100.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_NOTHING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_check_nothing .
  IF gt_alv IS INITIAL .
    PERFORM  frm_get_nothing.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_FIELDCAT
*&---------------------------------------------------------------------*
*       text  初始化ALV输出字段
*----------------------------------------------------------------------*
FORM  frm_init_fieldcat  CHANGING pt_fieldcat  TYPE lvc_t_fcat.
  PERFORM  frm_0100_init_fieldcat CHANGING pt_fieldcat.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CHECK_NOTHING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_0100_init_fieldcat  CHANGING pt_fieldcat  TYPE lvc_t_fcat.
  REFRESH:pt_fieldcat.

  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
      i_structure_name       = 'ZIDT_T_0003'
    CHANGING
      ct_fieldcat            = gt_fieldcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.

  mcr_fill_fieldcat:
    'BOX'         '选择',
    'ICON'        '状态',
    'MESG'        '消息'.

  LOOP AT gt_fieldcat ASSIGNING <gfs_fieldcat>.
    <gfs_fieldcat>-tabname = 'GT_ALV'.
    <gfs_fieldcat>-col_pos = sy-tabix.
    <gfs_fieldcat>-col_opt = 'A'.       "设置自动优化时，不生效

    CASE <gfs_fieldcat>-fieldname.
      WHEN 'BOX'.
        <gfs_fieldcat>-checkbox  = c_x.
        <gfs_fieldcat>-edit      = c_x.
        <gfs_fieldcat>-key       = c_x.
        <gfs_fieldcat>-col_opt   = space.
        <gfs_fieldcat>-col_pos   = 1.
        <gfs_fieldcat>-outputlen = '10'.
      WHEN 'ICON'.
        <gfs_fieldcat>-col_pos   = 2.
        <gfs_fieldcat>-key       = c_x.
      WHEN 'MESG'.
        <gfs_fieldcat>-col_pos   = 3.
        <gfs_fieldcat>-key       = c_x.
      WHEN 'INTEGRATE_GUID'.
        <gfs_fieldcat>-col_pos   = <gfs_fieldcat>-col_pos + 3.
        <gfs_fieldcat>-key       = space.
      WHEN 'IMPORT_DATAS' OR 'EXPORT_DATAS'.
        <gfs_fieldcat>-col_pos   = <gfs_fieldcat>-col_pos + 3.
        <gfs_fieldcat>-hotspot   = c_x.
      WHEN OTHERS.
        <gfs_fieldcat>-col_pos = <gfs_fieldcat>-col_pos + 3.
    ENDCASE.
  ENDLOOP.

  pt_fieldcat = gt_fieldcat.

ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_INIT_COMM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_init_comm .
  gs_layout-zebra      = c_x.
  gs_layout-cwidth_opt = c_x.
  gs_layout-sel_mode   = 'D'.
  gs_layout-no_rowmark = c_x.
  DESCRIBE TABLE gt_alv LINES DATA(lv_lines).
  gv_title = sy-title && '（共' && lv_lines && '条）'.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_STATUS_0100_PBO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM  frm_0100_status_pbo .
  DATA:lt_code LIKE TABLE OF sy-ucomm,
       ls_code LIKE sy-ucomm.
  RANGES: lr_ecode FOR sy-ucomm.

  mcr_range_eq: lr_ecode 'ALL'.
  mcr_range_eq: lr_ecode 'SAL'.

  DELETE lr_ecode WHERE low EQ 'ALL'.
  DELETE lr_ecode WHERE low EQ 'SAL'.

  LOOP AT lr_ecode.
    ls_code = lr_ecode-low.
    APPEND ls_code TO lt_code.
  ENDLOOP.

*设置标题与应用工具栏
  SET PF-STATUS 'PF_0100' EXCLUDING lt_code.
  SET TITLEBAR 'TI_0100' WITH gv_title.

*创建并划分ALV容器:前提ALV容易是初始化的
  PERFORM  frm_0100_init.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_CREATE_AND_INIT_CONTROLS
*&---------------------------------------------------------------------*
*       ALV容器处理
*----------------------------------------------------------------------*
FORM  frm_0100_init .
  DATA: lv_save,                         "for Parameter I_SAVE-------为变式保存参数
        ls_variant TYPE disvariant.      "for parameter IS_VARIANT---为变式保存参数
  DATA: lt_exclude TYPE ui_functions.    "工具栏排除
  DATA: lv_repid   TYPE sy-repid.        "当前程序号

  IF go_alv IS INITIAL.
*   //变式参数赋值
    lv_save = 'A'.
    PERFORM  frm_get_variant CHANGING ls_variant.
*   //创建ALV控制
    CREATE OBJECT go_alv
      EXPORTING
        i_parent = cl_gui_container=>screen0.
*   //fieldcat设置
    PERFORM  frm_exclude_tb_functions0 CHANGING lt_exclude.
    PERFORM  frm_exclude_tb_functions1 CHANGING lt_exclude.

    gs_layout-info_fname = 'ZCOLOR'.     " 行颜色字段
    gs_layout-ctab_fname = 'CELLCOLOR'.  " 单元格颜色字段
    gs_layout-stylefname = 'CELL_EDIT'.

    CALL METHOD go_alv->set_table_for_first_display
      EXPORTING
        is_variant           = ls_variant
        i_save               = lv_save
        is_layout            = gs_layout
        it_toolbar_excluding = lt_exclude
      CHANGING
        it_outtab            = gt_alv
        it_fieldcatalog      = gt_fieldcat.

    PERFORM  frm_alv_set_handler USING go_alv.
  ELSE.
    PERFORM  frm_refresh_000 USING go_alv.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_GET_VARIANT
*&---------------------------------------------------------------------*
*       ALV通用VARIANT设置
*----------------------------------------------------------------------*
FORM  frm_get_variant  CHANGING pc_variant TYPE disvariant.
  CONCATENATE sy-repid sy-dynnr INTO pc_variant-report.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_ALV_SET_HANDLER
*&---------------------------------------------------------------------*
*       text  OO ALV注册事件
*----------------------------------------------------------------------*
FORM  frm_alv_set_handler USING pa_grid TYPE REF TO cl_gui_alv_grid.
  IF zgo_handler IS INITIAL.
    CREATE OBJECT zgo_handler.
  ENDIF.

  SET HANDLER zgo_handler->handler_double_click FOR pa_grid.
  SET HANDLER zgo_handler->handle_hotspot_click FOR pa_grid.
  SET HANDLER zgo_handler->handle_user_command  FOR pa_grid.
  SET HANDLER zgo_handler->handler_data_changed FOR pa_grid.
  SET HANDLER zgo_handler->handle_toolbar       FOR pa_grid.
  SET HANDLER zgo_handler->onf4                 FOR pa_grid.
  PERFORM  frm_set_grid_onf4 USING   pa_grid.

*  CALL METHOD pa_grid->set_ready_for_input
*    EXPORTING
*      i_ready_for_input = 1.
  CALL METHOD pa_grid->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_enter.      "回车

  CALL METHOD pa_grid->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_modified.  "修改,既光标移出
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SET_GRID_ONF4
*&---------------------------------------------------------------------*
*       text  设置F4
*----------------------------------------------------------------------*
FORM  frm_set_grid_onf4 USING pa_grid TYPE REF TO cl_gui_alv_grid.
  DATA: lv_field TYPE lvc_fname,
        lt_f4    TYPE lvc_t_f4,
        ls_f4    TYPE lvc_s_f4.

  ls_f4-fieldname  = 'MTART'.
  ls_f4-register   = 'X'.
*  ls_f4-getbefore  = 'X'.
*  ls_f4-chngeafter = 'X'.
*  ls_f4-internal   = ''.
  INSERT ls_f4 INTO TABLE lt_f4.

  CALL METHOD pa_grid->register_f4_for_fields
    EXPORTING
      it_f4 = lt_f4.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_USER_COMMAND_0100_PAI
*&---------------------------------------------------------------------*
*       text  用户按键处理
*----------------------------------------------------------------------*
FORM  frm_0100_user_command_pai .
  DATA ls_mse TYPE string.
  gv_code = ok_code.
  CLEAR ok_code.
  CASE gv_code.
    WHEN 'ALL'.   " 全选
      PERFORM  frm_sel_data USING c_x.
    WHEN 'SAL'.
      PERFORM  frm_sel_data USING space.
    WHEN 'REFRESH'.  "
      PERFORM frm_uc_refresh.
    WHEN 'SEND'.  "
      PERFORM frm_uc_send.
  ENDCASE.
ENDFORM.
*&---------------------------------------------------------------------*
*&      FORM  FRM_SEL_DATA
*&---------------------------------------------------------------------*
*       text  全选/反选
*----------------------------------------------------------------------*
*  -->  pa_mark      text
*----------------------------------------------------------------------*
FORM  frm_sel_data USING pa_mark TYPE c.
  DATA:ls_alv LIKE gs_alv.
  ls_alv-box = pa_mark.
  MODIFY gt_alv FROM ls_alv TRANSPORTING box
  WHERE box NE pa_mark.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_UC_SAVE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_uc_send.
  LOOP AT gt_alv ASSIGNING FIELD-SYMBOL(<lfs_alv>)
    WHERE box  IS NOT INITIAL.
    CLEAR:
      <lfs_alv>-icon,
      <lfs_alv>-mesg.
    IF <lfs_alv>-import_datas IS INITIAL.
      <lfs_alv>-icon = icon_red_light.
      <lfs_alv>-mesg = '无导入报文，不能发送'.
    ENDIF.
  ENDLOOP.

  LOOP AT gt_alv ASSIGNING <lfs_alv>
    WHERE box  IS NOT INITIAL
      AND icon IS INITIAL.
    IF <lfs_alv>-access_sign = 'C'. "Consume Interface
      PERFORM frm_send_c CHANGING <lfs_alv> .
    ENDIF.

    IF <lfs_alv>-access_sign = 'S'."Supplier Interface
      PERFORM frm_send_s CHANGING <lfs_alv>.
    ENDIF.

    IF <lfs_alv>-send_guid IS NOT INITIAL.

      <lfs_alv>-icon = icon_green_light.
      <lfs_alv>-mesg = '发送成功'.

      UPDATE zidt_t_0003
         SET send_guid = <lfs_alv>-integrate_guid
       WHERE integrate_guid = <lfs_alv>-send_guid.
    ENDIF.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_send_c
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_send_c CHANGING cs_alv TYPE ty_alv.
  DATA:
    lv_user   TYPE ZIDT_E_USERNAME,
    lv_url    TYPE string,
    lv_token  TYPE string,
    ls_import TYPE zidt_s_import,
    ls_export TYPE zidt_s_export.

  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = cs_alv-integrate_guid.
  TRY .
      cs_alv-send_guid = cl_uuid_factory=>create_system_uuid( )->create_uuid_x16( ).
    CATCH cx_uuid_error.

  ENDTRY.

  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
   WHERE system_id    = @cs_alv-system_id
     AND integrate_id = @cs_alv-integrate_id.
  IF sy-subrc = 0.
    lv_url = ls_integrate-url.
*   需求TOKEN
    SEARCH lv_url FOR 'token='.
    IF sy-subrc = 0.
*      IF cs_alv-system_id = 'OA'.
*       SAP账户对应OA账户
*        SELECT SINGLE usr21~bname,usr21~persnumber,sort1,zbc_t_0001~loginname INTO @DATA(ls_user_info) FROM usr21
*          INNER JOIN adrp ON adrp~persnumber = usr21~persnumber
*          INNER JOIN zbc_t_0001 ON zbc_t_0001~code = adrp~sort1
*          WHERE usr21~bname = @cs_alv-integrate_user
*            AND zbc_t_0001~state = '1'    "在职
*            AND zbc_t_0001~isdeleted = @space. "删除状态为空
*      ENDIF.
*     无账号映射，使用SAP账户发送
*      IF ls_user_info IS INITIAL.
        lv_user = cs_alv-integrate_user.
*      ELSE.
*        lv_user = ls_user_info-loginname.
*      ENDIF.

      CALL METHOD zidt_c_integrate_tool=>get_token_from_outsyst
        EXPORTING
          iv_system_id       = cs_alv-system_id
          iv_user            = lv_user
        IMPORTING
          ev_token           = lv_token
        EXCEPTIONS
          no_token_configure = 1
          OTHERS             = 2.
      IF sy-subrc <> 0.
*         Implement suitable error handling here
      ENDIF.
      lv_url = lv_url && lv_token.
    ENDIF.
    MOVE-CORRESPONDING ls_integrate TO ls_import-import_parament-import_detail.
    ls_import-import_parament-integrate_guid = cs_alv-send_guid.
    ls_import-import_parament-import_detail-integrate_language = sy-langu.
    ls_import-import_parament-import_detail-integrate_user = sy-uname.
    ls_import-import_datas = ls_data-import_datas.
  ENDIF.

  IF lv_url IS NOT INITIAL.
*   写入导入记录
    CALL METHOD zidt_c_integrate_tool=>write_import_log
      EXPORTING
        is_import_parament = ls_import-import_parament
        iv_import_datas    = ls_import-import_datas.
*   发送HTTP
    CALL FUNCTION 'ZIDT_SET_HTTP'
      EXPORTING
        iv_url    = lv_url
        is_import = ls_import
      IMPORTING
        es_export = ls_export.

    ls_export-export_parament-reference_no = cs_alv-reference_no.
    ls_export-export_parament-sap_no = cs_alv-sap_no.
*   写入导出记录
    CALL METHOD zidt_c_integrate_tool=>write_export_log
      EXPORTING
        is_export_parament = ls_export-export_parament
        iv_export_datas    = ls_export-export_datas.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_send_s
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_send_s  CHANGING cs_alv TYPE ty_alv.
  DATA:
    ls_import  TYPE zidt_s_import,
    ls_export  TYPE zidt_s_export,
    lcl_import TYPE REF TO cl_abap_structdescr,
    lcr_import TYPE REF TO data,
    lcl_export TYPE REF TO cl_abap_structdescr,
    lcr_export TYPE REF TO data.

  READ TABLE gt_datas INTO DATA(ls_data) WITH KEY integrate_guid = cs_alv-integrate_guid.

  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
   WHERE system_id    = @cs_alv-system_id
     AND integrate_id = @cs_alv-integrate_id.

  MOVE-CORRESPONDING cs_alv TO ls_import-import_parament-import_detail.
  TRY .
      ls_import-import_parament-integrate_guid = cl_uuid_factory=>create_system_uuid( )->create_uuid_x16( ).
    CATCH cx_uuid_error.

  ENDTRY.
  cs_alv-send_guid = ls_import-import_parament-integrate_guid.
* 写入导入记录
  CALL METHOD zidt_c_integrate_tool=>write_import_log
    EXPORTING
      is_import_parament = ls_import-import_parament
      iv_import_datas    = ls_data-import_datas.
  MOVE-CORRESPONDING ls_import-import_parament TO ls_export-export_parament.
* 调用配置函数
  IF ls_export-export_parament-export_detail-integrate_type IS INITIAL.
    TRY .
        lcl_import ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_integrate-import_structure ).
        CREATE DATA lcr_import TYPE HANDLE lcl_import.
        ASSIGN lcr_import->* TO FIELD-SYMBOL(<lfs_import>).

        lcl_export ?= cl_abap_typedescr=>describe_by_name( EXPORTING p_name = ls_integrate-export_structure ).
        CREATE DATA lcr_export TYPE HANDLE lcl_export.
        ASSIGN lcr_export->* TO FIELD-SYMBOL(<lfs_export>).
      CATCH  cx_sy_move_cast_error.

    ENDTRY.
    IF <lfs_import> IS ASSIGNED
      AND <lfs_export> IS ASSIGNED.
*       转换JSON到结构话数据
      /ui2/cl_json=>deserialize( EXPORTING json        = ls_data-import_datas
                                           pretty_name = /ui2/cl_json=>pretty_mode-camel_case
                                 CHANGING  data        = <lfs_import> ).
      ASSIGN COMPONENT 'IMPORT_PARAMENT' OF STRUCTURE <lfs_import> TO FIELD-SYMBOL(<lfs_import_parament>).
      MOVE-CORRESPONDING ls_import-import_parament TO <lfs_import_parament>.
      CALL FUNCTION ls_integrate-function
        EXPORTING
          is_import = <lfs_import>
        IMPORTING
          es_export = <lfs_export>.
    ENDIF.
  ENDIF.

  IF <lfs_export> IS ASSIGNED.
    DATA(ev_json) = /ui2/cl_json=>serialize( data        = <lfs_export>
                                             pretty_name = /ui2/cl_json=>pretty_mode-none ).

    ASSIGN COMPONENT 'EXPORT_PARAMENT' OF STRUCTURE <lfs_export> TO FIELD-SYMBOL(<lfs_export_parament>).
    MOVE-CORRESPONDING <lfs_export_parament> TO ls_export-export_parament.
  ELSE.
    ev_json = /ui2/cl_json=>serialize( data        = ls_export
                                       pretty_name = /ui2/cl_json=>pretty_mode-none ).

  ENDIF.
  IF ls_export-export_parament-integrate_guid IS INITIAL.
    ls_export-export_parament-integrate_guid = ls_import-import_parament-integrate_guid.
  ENDIF.
* 写入导出记录
  CALL METHOD zidt_c_integrate_tool=>write_export_log
    EXPORTING
      is_export_parament = ls_export-export_parament
      iv_export_datas    = ev_json.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_get_token
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_get_token  USING    uv_system_id
                             uv_user
                    CHANGING cv_token.
*  DATA:
*    lv_iid       TYPE string,
*    ls_token_in  TYPE zoa_s0000_import,
*    ls_token_out TYPE zstoatokenr,
*    ls_import    TYPE zidt_s_import,
*    ls_export    TYPE zidt_s_export.
*  lv_iid = uv_system_id && '_TOKEN'.
*  SELECT SINGLE * INTO @DATA(ls_integrate) FROM zidt_t_0002
*   WHERE system_id    = @uv_system_id
*     AND integrate_id = @lv_iid.
** SAP账户对应OA账户
*  SELECT SINGLE usr21~bname,usr21~persnumber,sort1,zbc_t_0001~loginname INTO @DATA(ls_user_info) FROM usr21
*    INNER JOIN adrp ON adrp~persnumber = usr21~persnumber
*    INNER JOIN zbc_t_0001 ON zbc_t_0001~code = adrp~sort1
*    WHERE usr21~bname = @uv_user
*      AND zbc_t_0001~state = '1'    "在职
*      AND zbc_t_0001~isdeleted = @space. "删除状态为空
*
*  ls_token_in-user_name  = ls_integrate-username.
*  ls_token_in-password   = ls_integrate-password.
*  ls_token_in-login_name = ls_user_info-loginname.
*
*  DATA(lv_token_json) = /ui2/cl_json=>serialize( data        = ls_token_in
*                                                 compress    = abap_true
*                                                 pretty_name = /ui2/cl_json=>pretty_mode-camel_case ).
*
*  MOVE-CORRESPONDING ls_integrate TO ls_import-import_parament-import_detail.
*  ls_import-import_datas = lv_token_json.
*
** 发送HTTP
*  CALL FUNCTION 'ZIDT_SET_HTTP'
*    EXPORTING
*      is_import = ls_import
*    IMPORTING
*      es_export = ls_export.
*  IF ls_export-export_datas IS NOT INITIAL.
*    /ui2/cl_json=>deserialize( EXPORTING json        = ls_export-export_datas
*                                         pretty_name = /ui2/cl_json=>pretty_mode-low_case
*                               CHANGING  data        = ls_token_out ).
*  ENDIF.
*  cv_token = ls_token_out-id.
*  IF cv_token IS INITIAL.
*    MESSAGE e208(00) WITH TEXT-m01.
*  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form frm_uc_refresh
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM frm_uc_refresh .

  PERFORM frm_data_get.

  PERFORM frm_data_deal.
  CLEAR gv_title.
  DESCRIBE TABLE gt_alv LINES DATA(lv_lines).
  gv_title = sy-title+0(9) && '（共' && lv_lines && '条）'.
ENDFORM.