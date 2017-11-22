    PACKAGE excel  
    IS  
       /*  
                 Global excel.Application Object --> this represent excel Object.  
         */  
       appl_id   client_ole2.obj_type;  
      
       /*  
               Open file that act as template. Parameter are:  
               _application_ -- global word parameter that we initialize at   
               begining.  
               _file_ -- file name we wish to open --> it can be from database, or filesystem...  
       */  
       FUNCTION file_open (application client_ole2.obj_type, FILE VARCHAR2)  
          RETURN client_ole2.obj_type;  
      
       /*  
               Close current file.  
       */  
       PROCEDURE file_close (document client_ole2.obj_type);  
      
       /*  
               Saves current file (It is useful if we need to save current   
               file using another name)  
       */  
       PROCEDURE file_save_as (document client_ole2.obj_type, FILE VARCHAR2);  
      
       /*  
               Isert number (not formated)  
               x - horizontal axei.  
               y - vertical axis.  
               v - value.  
       */  
       PROCEDURE insert_number (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           NUMBER  
       );  
      
       /*  
               Insert number and format it as decimal value.   
               x - horizontal axei.  
               y - vertical axis.  
               v - value.  
               Napomena: !!!THIS DOES NOT WORK IN EXCEL 2007!!!  
       */  
       PROCEDURE insert_number_decimal (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           NUMBER  
       );  
      
       /*  
               Insert characters  (not formated)  
               x - horizontal axei.  
               y - vertical axis.  
               v - value.  
       */  
       PROCEDURE insert_char (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           VARCHAR2  
       );  
      
       /*  
               Insert character - formated  
               color - numbers (15 for example is gray)             
               style - BOLD' or 'ITALIC'  
               x - horizontal axei.  
               y - vertical axis.  
               v - value.  
       */  
       PROCEDURE insert_char_formated (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           VARCHAR2,  
          color       NUMBER,  
          style       VARCHAR2  
       );  
      
       /*  
               Set autofit on whole sheet.  
       */  
       PROCEDURE set_auto_fit (worksheet client_ole2.obj_type);  
      
       /*  
               Set autofit for range r. For example. r can be: 'A2:E11'  
       */  
       PROCEDURE set_auto_fit_range (worksheet client_ole2.obj_type, r VARCHAR2);  
      
       /*  
               Put decimal format (0.00) on range r.  
       */  
       PROCEDURE set_decimal_format_range (  
          worksheet   client_ole2.obj_type,  
          r           VARCHAR2  
       );  
      
       /*  
               Create new workbook.  
       */  
       FUNCTION new_workbook (application client_ole2.obj_type)  
          RETURN client_ole2.obj_type;  
      
       /*  
               Create new worksheet.  
       */  
       FUNCTION new_worksheet (workbook client_ole2.obj_type)  
          RETURN client_ole2.obj_type;  
      
       /*  
               Saves file in client tempfolder (It is necessary to save file if edit template).  
       */  
       FUNCTION download_file (  
          file_name         IN   VARCHAR2,  
          table_name        IN   VARCHAR2,  
          column_name       IN   VARCHAR2,  
          where_condition   IN   VARCHAR2  
       )  
          RETURN VARCHAR2;  
      
       /*  
               Run macro on client excel document.  
       */  
       PROCEDURE run_macro_on_document (  
          document   client_ole2.obj_type,  
          macro      VARCHAR2  
       );  
      
       /*  
               Limit network load...not important.  
       */  
       PROCEDURE insert_number_array (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           VARCHAR2  
       );  
    END;  

Package body:
view plainprint?

    PACKAGE BODY excel  
    IS  
       FUNCTION file_open (application client_ole2.obj_type, FILE VARCHAR2)  
          RETURN client_ole2.obj_type  
       IS  
          arg_list    client_ole2.list_type;  
          document    client_ole2.obj_type;  
          documents   client_ole2.obj_type;  
       BEGIN  
          arg_list := client_ole2.create_arglist;  
          documents := client_ole2.invoke_obj (application, 'Workbooks');  
          client_ole2.add_arg (arg_list, FILE);  
          document := client_ole2.invoke_obj (documents, 'Open', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (documents);  
          RETURN document;  
       END file_open;  
      
       PROCEDURE file_save_as (document client_ole2.obj_type, FILE VARCHAR2)  
       IS  
          arg_list   client_ole2.list_type;  
       BEGIN  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, FILE);  
          client_ole2.invoke (document, 'SaveAs', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
       END file_save_as;  
      
       FUNCTION new_workbook (application client_ole2.obj_type)  
          RETURN client_ole2.obj_type  
       IS  
          workbook    client_ole2.obj_type;  
          workbooks   client_ole2.obj_type;  
       BEGIN  
          workbooks := client_ole2.get_obj_property (application, 'Workbooks');  
          workbook := client_ole2.invoke_obj (workbooks, 'Add');  
          client_ole2.RELEASE_OBJ (workbooks);  
          RETURN workbook;  
       END new_workbook;  
      
       FUNCTION new_worksheet (workbook client_ole2.obj_type)  
          RETURN client_ole2.obj_type  
       IS  
          worksheets   client_ole2.obj_type;  
          worksheet    client_ole2.obj_type;  
       BEGIN  
          worksheets := client_ole2.get_obj_property (workbook, 'Worksheets');  
          worksheet := client_ole2.invoke_obj (worksheets, 'Add');  
          client_ole2.RELEASE_OBJ (worksheets);  
          RETURN worksheet;  
       END new_worksheet;  
      
       PROCEDURE file_close (document client_ole2.obj_type)  
       IS  
       BEGIN  
          client_ole2.invoke (document, 'Close');  
       END file_close;  
      
       /*  
           Macro:    Cells(3, 4).Value = 3  
                       Cells(3, 4).Select  
                       Selection.NumberFormat = "0.00"  
       */  
       PROCEDURE insert_number_decimal (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           NUMBER  
       )  
       IS  
          args        client_ole2.list_type;  
          cell        client_ole2.obj_type;  
          selection   client_ole2.obj_type;  
       BEGIN  
          IF v IS NOT NULL  
          THEN  
             args := client_ole2.create_arglist;  
             client_ole2.add_arg (args, x);  
             client_ole2.add_arg (args, y);  
             cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
             client_ole2.destroy_arglist (args);  
             client_ole2.set_property (cell, 'Value', v);  
             client_ole2.invoke (cell, 'Select');  
             selection := client_ole2.invoke_obj (appl_id, 'Selection');  
             client_ole2.set_property (selection, 'Numberformat', '#.##0,00');  
             client_ole2.RELEASE_OBJ (selection);  
             client_ole2.RELEASE_OBJ (cell);  
          END IF;  
       END;  
      
       /* Macro:  
                           Cells(x, y).Value = v  
       */  
       PROCEDURE insert_number (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           NUMBER  
       )  
       IS  
          args   client_ole2.list_type;  
          cell   ole2.obj_type;  
       BEGIN  
          IF v IS NOT NULL  
          THEN  
             args := client_ole2.create_arglist;  
             client_ole2.add_arg (args, x);  
             client_ole2.add_arg (args, y);  
             cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
             client_ole2.destroy_arglist (args);  
             client_ole2.set_property (cell, 'Value', v);  
             client_ole2.RELEASE_OBJ (cell);  
          END IF;  
       END;  
      
      
       PROCEDURE insert_char (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           VARCHAR2  
       )  
       IS  
          args   client_ole2.list_type;  
          cell   client_ole2.obj_type;  
       BEGIN  
          IF v IS NOT NULL  
          THEN  
             args := client_ole2.create_arglist;  
             client_ole2.add_arg (args, x);  
             client_ole2.add_arg (args, y);  
             cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
             client_ole2.destroy_arglist (args);  
             client_ole2.set_property (cell, 'Value', v);  
             client_ole2.RELEASE_OBJ (cell);  
          END IF;  
       END;  
      
        
       /*  
               Macro:  
                           Cells(x, y).Value = v  
                           Cells(x, y).Select  
                           Selection.Interior.ColorIndex = color  
                           if (style in 'BOLD')  
                               Selection.Font.Bold = True  
                           else if (style in 'ITALIC')  
                               Selection.Font.Italic = True  
       */  
       PROCEDURE insert_char_formated (  
          worksheet   client_ole2.obj_type,  
          x           NUMBER,  
          y           NUMBER,  
          v           VARCHAR2,  
          color       NUMBER,  
          style       VARCHAR2  
       )  
       IS  
          args        client_ole2.list_type;  
          cell        client_ole2.obj_type;  
          selection   client_ole2.obj_type;  
          font        client_ole2.obj_type;  
          interior    client_ole2.obj_type;  
       BEGIN  
          IF v IS NOT NULL  
          THEN  
             args := client_ole2.create_arglist;  
             client_ole2.add_arg (args, x);  
             client_ole2.add_arg (args, y);  
             cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
             client_ole2.destroy_arglist (args);  
             client_ole2.set_property (cell, 'Value', v);  
             client_ole2.invoke (cell, 'Select');  
             selection := client_ole2.invoke_obj (appl_id, 'Selection');  
             font := client_ole2.invoke_obj (selection, 'Font');  
             interior := client_ole2.invoke_obj (selection, 'Interior');  
      
             IF UPPER (style) IN ('BOLD', 'ITALIC')  
             THEN  
                client_ole2.set_property (font, style, TRUE);  
             END IF;  
      
             client_ole2.set_property (interior, 'ColorIndex', color);  
             client_ole2.RELEASE_OBJ (interior);  
             client_ole2.RELEASE_OBJ (font);  
             client_ole2.RELEASE_OBJ (selection);  
             client_ole2.RELEASE_OBJ (cell);  
          END IF;  
       END;  
      
       /*  
               Macro:  
                           Range(r).Select  
                           Selection.Columns.AutoFit  
                           Cells(1,1).Select  
       */  
       PROCEDURE set_auto_fit_range (worksheet client_ole2.obj_type, r VARCHAR2)  
       IS  
          args        client_ole2.list_type;  
          --range  
          rang        client_ole2.obj_type;  
          selection   client_ole2.obj_type;  
          colum       client_ole2.obj_type;  
          cell        client_ole2.obj_type;  
       BEGIN  
          args := client_ole2.create_arglist;  
          client_ole2.add_arg (args, r);  
          rang := client_ole2.get_obj_property (worksheet, 'Range', args);  
          client_ole2.destroy_arglist (args);  
          client_ole2.invoke (rang, 'Select');  
          selection := client_ole2.invoke_obj (appl_id, 'Selection');  
          colum := client_ole2.invoke_obj (selection, 'Columns');  
          client_ole2.invoke (colum, 'AutoFit');  
          --now select upper (1,1) for deselection.        
          args := client_ole2.create_arglist;  
          client_ole2.add_arg (args, 1);  
          client_ole2.add_arg (args, 1);  
          cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
          client_ole2.invoke (cell, 'Select');  
          client_ole2.destroy_arglist (args);  
          client_ole2.RELEASE_OBJ (colum);  
          client_ole2.RELEASE_OBJ (selection);  
          client_ole2.RELEASE_OBJ (rang);  
       END set_auto_fit_range;  
      
       /*  
               Macro:  
                           Range(r).Select  
                           Selection.Numberformat = "0.00"  
                           Cells(1,1).Select  
       */  
       PROCEDURE set_decimal_format_range (  
          worksheet   client_ole2.obj_type,  
          r           VARCHAR2  
       )  
       IS  
          args        client_ole2.list_type;  
          --range  
          rang        client_ole2.obj_type;  
          selection   client_ole2.obj_type;  
          --colum Client_OLE2.Obj_Type;  
          cell        client_ole2.obj_type;  
       BEGIN  
          args := client_ole2.create_arglist;  
          client_ole2.add_arg (args, r);  
          rang := client_ole2.get_obj_property (worksheet, 'Range', args);  
          client_ole2.destroy_arglist (args);  
          client_ole2.invoke (rang, 'Select');  
          selection := client_ole2.invoke_obj (appl_id, 'Selection');  
          --colum:= Client_OLE2.invoke_obj(selection, 'Columns');  
          client_ole2.set_property (selection, 'Numberformat', '#.##0,00');  
          --Client_OLE2.invoke(colum, 'AutoFit');  
          --now select upper (1,1) for deselection.        
          args := client_ole2.create_arglist;  
          client_ole2.add_arg (args, 1);  
          client_ole2.add_arg (args, 1);  
          cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
          client_ole2.invoke (cell, 'Select');  
          client_ole2.destroy_arglist (args);  
          --Client_OLE2.release_obj(colum);  
          client_ole2.RELEASE_OBJ (selection);  
          client_ole2.RELEASE_OBJ (rang);  
       END set_decimal_format_range;  
      
       /*  
               Macro:Cells.Select  
                           Selection.Columns.AutoFit  
                           Cells(1,1).Select  
       */  
       PROCEDURE set_auto_fit (worksheet client_ole2.obj_type)  
       IS  
          args        client_ole2.list_type;  
          cell        client_ole2.obj_type;  
          selection   client_ole2.obj_type;  
          colum       client_ole2.obj_type;  
       BEGIN  
          cell := client_ole2.get_obj_property (worksheet, 'Cells');  
          client_ole2.invoke (cell, 'Select');  
          selection := client_ole2.invoke_obj (appl_id, 'Selection');  
          colum := client_ole2.invoke_obj (selection, 'Columns');  
          client_ole2.invoke (colum, 'AutoFit');  
          --now select upper (1,1) for deselection.        
          args := client_ole2.create_arglist;  
          client_ole2.add_arg (args, 1);  
          client_ole2.add_arg (args, 1);  
          cell := client_ole2.get_obj_property (worksheet, 'Cells', args);  
          client_ole2.invoke (cell, 'Select');  
          client_ole2.destroy_arglist (args);  
          client_ole2.RELEASE_OBJ (colum);  
          client_ole2.RELEASE_OBJ (selection);  
          client_ole2.RELEASE_OBJ (cell);  
       END set_auto_fit;  
      
       PROCEDURE run_macro_on_document (  
          document   client_ole2.obj_type,  
          macro      VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, macro);  
          client_ole2.invoke (excel.appl_id, 'Run', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
       END;  
      
       FUNCTION download_file (  
          file_name         IN   VARCHAR2,  
          table_name        IN   VARCHAR2,  
          column_name       IN   VARCHAR2,  
          where_condition   IN   VARCHAR2  
       )  
          RETURN VARCHAR2  
       IS  
          l_ok          BOOLEAN;  
          c_file_name   VARCHAR2 (512);  
          c_path        VARCHAR2 (255);  
       BEGIN  
          SYNCHRONIZE;  
          c_path := client_win_api_environment.get_temp_directory (FALSE);  
      
          IF c_path IS NULL  
          THEN  
             c_path := 'C:\';  
          ELSE  
             c_path := c_path || '\';  
          END IF;  
      
          c_file_name := c_path || file_name;  
          l_ok :=  
             webutil_file_transfer.db_to_client_with_progress  
                                                       (c_file_name,  
                                                        table_name,  
                                                        column_name,  
                                                        where_condition,  
                                                        'Transfer on file system',  
                                                        'Progress'  
                                                       );  
          SYNCHRONIZE;  
      
          IF NOT l_ok  
          THEN  
             msg_popup ('File not found in database', 'E', TRUE);  
          END IF;  
      
          RETURN c_path || file_name;  
       END download_file;  
    END;  