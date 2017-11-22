    PACKAGE word  
    IS  
       /*  
               Global Word.Application Object --> represent word object.  
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
               (Bizniss end of this whole package;) ) Inserts value in specific word bookmark.  
               _dokcument_ -- Word document.  
               _bookmark_ -- Name of bookmark that is defined in word template,  
               _content_ --  Content we wish to insert into bookmark.  
       */  
       PROCEDURE insertafter_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       );  
      
       /*  
               InsertAfter_Bookmark insert after bookmark and then delete that bookmark and this is not  
               good if you itarate through values, so this one do not delete bookmark after insert.  
               same paramters as previous one.  
       */  
       PROCEDURE replace_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       );  
      
       /*  
               Saame as previous procedure but it handle next for you.  
       */  
       PROCEDURE insertafter_bookmark_next (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       );  
      
       /*  
               This one after value insert move itself on next row into table. When I say next I mean next-down.  
               This is essential for iterating through word table (one row at the time)  
               We need manualy create new row if it does not existsexists.!!!  
       */  
       PROCEDURE insertafter_bookmark_down (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       );  
      
       /*  
               Easy...delete bookmark,  
       */  
       PROCEDURE delete_bookmark (document client_ole2.obj_type, bookmark VARCHAR2);  
      
       /*  
               Create new table row (see InsertAfter_Bookmark_Next)  
       */  
       PROCEDURE insert_new_table_row (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2  
       );  
      
       /*  
               Move bookmakr (ONLY IN TABLE) left, right, up, down.  
               _direction_ can have following valyes'UP', 'DOWN', 'LEFT', 'RIGHT'  
       */  
       PROCEDURE move_table_bookmark (  
          document    client_ole2.obj_type,  
          bookmark    VARCHAR2,  
          direction   VARCHAR2  
       );  
      
       /*  
               File download.  
               parametar _file_name_  -- client file name (name on client)  
               _table_name_ -- Table name for where BLOB column is.  
               _column_name_ -- BLOB column name that holds Word template.  
               -where_condition_ -- filter.  
       */  
       FUNCTION download_file (  
          file_name         IN   VARCHAR2,  
          table_name        IN   VARCHAR2,  
          column_name       IN   VARCHAR2,  
          where_condition   IN   VARCHAR2  
       )  
          RETURN VARCHAR2;  
      
       /*  
               Calling macro's on bookmarks...only for test.  
       */  
       PROCEDURE run_macro_on_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          macro      VARCHAR2  
       );  
      
       PROCEDURE run_macro_on_document (  
          document   client_ole2.obj_type,  
          macro      VARCHAR2  
       );  
    END;  

/*************************************************************/
    PACKAGE BODY word  
    IS  
       FUNCTION file_open (application client_ole2.obj_type, FILE VARCHAR2)  
          RETURN client_ole2.obj_type  
       IS  
          arg_list    client_ole2.list_type;  
          document    client_ole2.obj_type;  
          documents   client_ole2.obj_type;  
       BEGIN  
          arg_list := client_ole2.create_arglist;  
          documents := client_ole2.invoke_obj (application, 'documents');  
          client_ole2.add_arg (arg_list, FILE);  
          document := client_ole2.invoke_obj (documents, 'Open', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (documents);  
          RETURN document;  
       END file_open;  
      
       PROCEDURE file_close (document client_ole2.obj_type)  
       IS  
       BEGIN  
          client_ole2.invoke (document, 'Close');  
       --CLIENT_OLE2.RELEASE_OBJ(document);  
       END file_close;  
      
       PROCEDURE file_save_as (document client_ole2.obj_type, FILE VARCHAR2)  
       IS  
          arg_list   client_ole2.list_type;  
       BEGIN  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, FILE);  
          client_ole2.invoke (document, 'SaveAs', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
       --CLIENT_OLE2.RELEASE_OBJ(document);  
       END file_save_as;  
      
       PROCEDURE replace_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, content);  
          client_ole2.invoke (selectionobj, 'Delete');  
          client_ole2.invoke (selectionobj, 'InsertAfter', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END replace_bookmark;  
      
       PROCEDURE insertafter_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, content);  
          client_ole2.invoke (selectionobj, 'InsertAfter', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END insertafter_bookmark;  
      
       PROCEDURE insertafter_bookmark_next (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, content || CHR (13));  
          client_ole2.invoke (selectionobj, 'InsertAfter', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END insertafter_bookmark_next;  
      
       PROCEDURE insertafter_bookmark_down (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          content    VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, content);  
          client_ole2.invoke (selectionobj, 'InsertAfter', arg_list);  
          client_ole2.invoke (selectionobj, 'Cut');  
          client_ole2.invoke (selectionobj, 'SelectCell');  
          client_ole2.invoke (selectionobj, 'MoveDown');  
          client_ole2.invoke (selectionobj, 'Paste');  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END insertafter_bookmark_down;  
      
       PROCEDURE delete_bookmark (document client_ole2.obj_type, bookmark VARCHAR2)  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          client_ole2.invoke (selectionobj, 'Delete');  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END delete_bookmark;  
      
       PROCEDURE run_macro_on_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2,  
          macro      VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, macro);  
          client_ole2.invoke (word.appl_id, 'Run', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END;  
      
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
          --bookmarkCollection := CLIENT_OLE2.INVOKE_OBJ(document, 'Bookmarks', arg_list);  
          --arg_list := CLIENT_OLE2.CREATE_ARGLIST;  
          --CLIENT_OLE2.ADD_ARG(arg_list, bookmark);  
          --bookmarkObj := CLIENT_OLE2.INVOKE_OBJ(bookmarkCollection, 'Item',arg_list);  
          --CLIENT_OLE2.DESTROY_ARGLIST(arg_list);  
      
          --CLIENT_OLE2.INVOKE(bookmarkObj, 'Select');  
          --selectionObj := CLIENT_OLE2.INVOKE_OBJ(appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, macro);  
          client_ole2.invoke (word.appl_id, 'Run', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
       --CLIENT_OLE2.RELEASE_OBJ(selectionObj);  
       --CLIENT_OLE2.RELEASE_OBJ(bookmarkObj);  
       --CLIENT_OLE2.RELEASE_OBJ(bookmarkCollection);  
       END;  
      
       PROCEDURE insert_new_table_row (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, 1);  
          client_ole2.invoke (selectionobj, 'InsertRowsBelow', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END insert_new_table_row;  
      
       PROCEDURE move_down_table_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          client_ole2.invoke (selectionobj, 'Cut');  
          client_ole2.invoke (selectionobj, 'SelectCell');  
          client_ole2.invoke (selectionobj, 'MoveDown');  
          client_ole2.invoke (selectionobj, 'Paste');  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END move_down_table_bookmark;  
      
       PROCEDURE move_up_table_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          client_ole2.invoke (selectionobj, 'Cut');  
          client_ole2.invoke (selectionobj, 'SelectCell');  
          client_ole2.invoke (selectionobj, 'MoveUp');  
          client_ole2.invoke (selectionobj, 'Paste');  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END move_up_table_bookmark;  
      
       PROCEDURE move_left_table_bookmark (  
          document   client_ole2.obj_type,  
          bookmark   VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
          client_ole2.invoke (selectionobj, 'Cut');  
          client_ole2.invoke (selectionobj, 'SelectCell');  
          client_ole2.invoke (selectionobj, 'MoveUp');  
          client_ole2.invoke (selectionobj, 'Paste');  
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END move_left_table_bookmark;  
      
       PROCEDURE move_table_bookmark (  
          document    client_ole2.obj_type,  
          bookmark    VARCHAR2,  
          direction   VARCHAR2  
       )  
       IS  
          arg_list             client_ole2.list_type;  
          bookmarkcollection   client_ole2.obj_type;  
          bookmarkobj          client_ole2.obj_type;  
          selectionobj         client_ole2.obj_type;  
       BEGIN  
          bookmarkcollection :=  
                         client_ole2.invoke_obj (document, 'Bookmarks', arg_list);  
          arg_list := client_ole2.create_arglist;  
          client_ole2.add_arg (arg_list, bookmark);  
          bookmarkobj :=  
                    client_ole2.invoke_obj (bookmarkcollection, 'Item', arg_list);  
          client_ole2.destroy_arglist (arg_list);  
          client_ole2.invoke (bookmarkobj, 'Select');  
          selectionobj := client_ole2.invoke_obj (appl_id, 'Selection');  
      
          IF UPPER (direction) IN ('UP', 'DOWN', 'LEFT', 'RIGHT')  
          THEN  
             client_ole2.invoke (selectionobj, 'Cut');  
             client_ole2.invoke (selectionobj, 'SelectCell');  
             client_ole2.invoke (selectionobj, 'Move' || direction);  
             client_ole2.invoke (selectionobj, 'Paste');  
          END IF;  
      
          client_ole2.RELEASE_OBJ (selectionobj);  
          client_ole2.RELEASE_OBJ (bookmarkobj);  
          client_ole2.RELEASE_OBJ (bookmarkcollection);  
       END move_table_bookmark;  
      
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
          RETURN c_path || file_name;  
       END download_file;  
    END;  