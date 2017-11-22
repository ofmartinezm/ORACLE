http://www.tek-tips.com/viewthread.cfm?qid=319377

DECLARE
   hApplication      CLIENT_OLE2.OBJ_TYPE;
   hDocuments        CLIENT_OLE2.OBJ_TYPE;
   hDocument         CLIENT_OLE2.OBJ_TYPE;
   hSelection        CLIENT_OLE2.OBJ_TYPE;
   hRange            CLIENT_OLE2.OBJ_TYPE;
   hTables           CLIENT_OLE2.OBJ_TYPE;
   hTable            CLIENT_OLE2.OBJ_TYPE;
   hCells            CLIENT_OLE2.OBJ_TYPE;
   hRows             CLIENT_OLE2.OBJ_TYPE;
   hShading          CLIENT_OLE2.OBJ_TYPE;
   hParagraphFormat  CLIENT_OLE2.OBJ_TYPE;
   hFont             CLIENT_OLE2.OBJ_TYPE;
   
   args              CLIENT_OLE2.LIST_TYPE;       
   
   ---- wdDefaultTableBehavior Class members
   wdWord8TableBehavior      CONSTANT number(5) :=  0;  --Default
   wdWord9TableBehavior      CONSTANT number(5) :=  1;
   
   ---- wdAutoFitBehavior Class members
   ---- (only works when DefaultTableBehavior = wdWord9TableBehavior)
   wdAutoFitContent          CONSTANT number(5) :=  1;
   wdAutoFitFixed            CONSTANT number(5) :=  0;
   wdAutoFitWindow           CONSTANT number(5) :=  2;
   
   ---- wdUnits Class members  
   wdCell                    CONSTANT number(5) := 12;
   wdCharacter               CONSTANT number(5) :=  1;   --Default
   wdWord                    CONSTANT number(5) :=  2;
   wdSentence                CONSTANT number(5) :=  3;
   wdLine                    CONSTANT number(5) :=  5;
   
   ---- wdMovementType Class members
   wdExtend                  CONSTANT number(5) :=  1;
   wdMove                    CONSTANT number(5) :=  0;   --Default
   
   ---- WdParagraphAlignment Class members
   wdAlignParagraphCenter    CONSTANT number(5) :=  1;
   wdAlignParagraphLeft      CONSTANT number(5) :=  0;
   wdAlignParagraphRight     CONSTANT number(5) :=  2;
   
   ---- HexColor = BBGGRR
   myLightBlue               CONSTANT number(8) := 16755370; --FFAAAA
   
   CURSOR c IS
      SELECT   CODALF product, VERSION_PRG amount, rownum
      FROM     TUIB_MODUL
      WHERE    rownum < 100;
   
BEGIN
   hApplication := CLIENT_OLE2.CREATE_OBJ('Word.Application');
   CLIENT_OLE2.SET_PROPERTY(hApplication, 'Visible', 1);
   
   hDocuments := CLIENT_OLE2.GET_OBJ_PROPERTY(hApplication, 'Documents');   
   --cLIENT_OLE2.add_arg(args, 'C:\temp\MYDOC.docx');
   hDocument  := CLIENT_OLE2.INVOKE_OBJ(hDocuments,'Add');                  
   
   
   
   


   ------------------------------
   -------- Create Table --------
   ------------------------------
   hSelection := CLIENT_OLE2.GET_OBJ_PROPERTY(hApplication, 'Selection');    
   hTables := CLIENT_OLE2.GET_OBJ_PROPERTY(hDocument , 'Tables' );
   hRange  := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'Range');    
   args := CLIENT_OLE2.CREATE_ARGLIST;   
   CLIENT_OLE2.ADD_ARG_OBJ(args, hRange);            --Range
   CLIENT_OLE2.ADD_ARG(args, 3);                     --NumRows
   CLIENT_OLE2.ADD_ARG(args, 2);                     --NumColumns
   CLIENT_OLE2.ADD_ARG(args, wdWord9TableBehavior);  --DefaultTableBehavior
   CLIENT_OLE2.ADD_ARG(args, wdAutoFitContent);      --AutoFitBehavior
   hTable := CLIENT_OLE2.INVOKE_OBJ(hTables, 'Add', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   CLIENT_OLE2.RELEASE_OBJ(hTable);
   CLIENT_OLE2.RELEASE_OBJ(hRange);
   CLIENT_OLE2.RELEASE_OBJ(hTables);
    
   ------------------------------------------
   -------- Create and Format Header --------
   ------------------------------------------
   
   ---- Add next 2 cells to the selection
   ---- (next 2 cells are actually row 2)
   AG_MESSAGE('next 2 cells are actually row 2','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdCharacter);           --Unit
   CLIENT_OLE2.ADD_ARG(args, 2);                     --Count
   CLIENT_OLE2.ADD_ARG(args, wdExtend);              --Extend
   CLIENT_OLE2.INVOKE(hSelection, 'MoveRight', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);    
   
   ---- Set color of selected cells
   AG_MESSAGE('Set color of selected cells','I');
   hcells := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'Cells');    
   hShading := CLIENT_OLE2.GET_OBJ_PROPERTY(hCells,   'Shading');    
   CLIENT_OLE2.SET_PROPERTY(hShading, 'BackgroundPatternColor', myLightBlue);
   CLIENT_OLE2.RELEASE_OBJ(hShading);
   CLIENT_OLE2.RELEASE_OBJ(hCells);
   
   ---- Set selected cells to be Bold Header
   /*AG_MESSAGE('Set selected cells to be Bold Header','I');
   hrows := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'Rows');    
   CLIENT_OLE2.SET_PROPERTY(hRows, 'AllowBreakAcrossPages', True);
   CLIENT_OLE2.SET_PROPERTY(hRows, 'HeadingFormat', True);
   CLIENT_OLE2.RELEASE_OBJ(hRows);
   hFont := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'Font');    
   CLIENT_OLE2.SET_PROPERTY(hFont, 'Bold', True);
   CLIENT_OLE2.RELEASE_OBJ(hFont);*/
   
   ---- Move to Header row 1, set text and center
   AG_MESSAGE(' Move to Header row 1, set text and center','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdCharacter);
   CLIENT_OLE2.ADD_ARG(args, 1);
   CLIENT_OLE2.ADD_ARG(args, wdMove);
   CLIENT_OLE2.INVOKE(hSelection, 'MoveLeft', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   hParagraphFormat := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'ParagraphFormat');    
   CLIENT_OLE2.SET_PROPERTY(hParagraphFormat, 'Alignment', wdAlignParagraphCenter);
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, 'Sales');
   CLIENT_OLE2.INVOKE(hSelection, 'TypeText', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   ---- Move to Header row 2/cell 1, set text
   AG_MESSAGE(' Move to Header row 2/cell 1, set text','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdCell);
   CLIENT_OLE2.ADD_ARG(args, 1);
   CLIENT_OLE2.ADD_ARG(args, wdMove);
   CLIENT_OLE2.INVOKE(hSelection, 'MoveRight', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, 'Product');
   CLIENT_OLE2.INVOKE(hSelection, 'TypeText', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   ---- Move to Header row 2/cell 2, set text and right align
   AG_MESSAGE('Move to Header row 2/cell 2, set text and right align','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdCell);
   CLIENT_OLE2.ADD_ARG(args, 1);
   CLIENT_OLE2.ADD_ARG(args, wdMove);
   CLIENT_OLE2.INVOKE(hSelection, 'MoveRight', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   hParagraphFormat := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'ParagraphFormat');    
   CLIENT_OLE2.SET_PROPERTY(hParagraphFormat, 'Alignment', wdAlignParagraphRight);
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, 'Amount');
   CLIENT_OLE2.INVOKE(hSelection, 'TypeText', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   ---- Move row 3/cell 2, right align
   AG_MESSAGE(' Move row 3/cell 2, right align','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdLine);
   CLIENT_OLE2.ADD_ARG(args, 1);
   CLIENT_OLE2.ADD_ARG(args, wdMove);
   CLIENT_OLE2.INVOKE(hSelection, 'MoveDown', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   hParagraphFormat := CLIENT_OLE2.GET_OBJ_PROPERTY(hSelection, 'ParagraphFormat');    
   CLIENT_OLE2.SET_PROPERTY(hParagraphFormat, 'Alignment', wdAlignParagraphRight);
   
   ---- Move back to first cell
   AG_MESSAGE('Move back to first cell','I');
   args := CLIENT_OLE2.CREATE_ARGLIST;
   CLIENT_OLE2.ADD_ARG(args, wdCell);
   CLIENT_OLE2.ADD_ARG(args, 1);
   CLIENT_OLE2.ADD_ARG(args, wdMove);
   CLIENT_OLE2.INVOKE(hSelection, 'MoveLeft', args);
   CLIENT_OLE2.DESTROY_ARGLIST(args);
   
   -------------------------------------------
   -------- Populate rows from cursor --------
   -------------------------------------------
   FOR v IN c
   LOOP
   	--AG_MESSAGE('Populate rows from cursor','I');
      /*IF v.rownum != 1 THEN
         ---- create new line by moving over 1 cell
         args := CLIENT_OLE2.CREATE_ARGLIST;
         CLIENT_OLE2.ADD_ARG(args, wdCell);
         CLIENT_OLE2.ADD_ARG(args, 1);
         CLIENT_OLE2.ADD_ARG(args, wdMove);
         CLIENT_OLE2.INVOKE(hSelection, 'MoveRight', args);
         CLIENT_OLE2.DESTROY_ARGLIST(args);
      END IF;*/
      
      ---- Print product, move over a cell
     -- AG_MESSAGE('Print product, move over a cell'||v.product,'I');
      CLIENT_OLE2.set_property (hSelection, 'Text',v.product );
      args := CLIENT_OLE2.CREATE_ARGLIST;
     	CLIENT_OLE2.add_arg (args, 12);
			CLIENT_OLE2.invoke (hSelection, 'MoveRight', args);
			CLIENT_OLE2.destroy_arglist (args);

      
           
      /*
      args := CLIENT_OLE2.CREATE_ARGLIST;
      CLIENT_OLE2.ADD_ARG(args, wdCell);
      CLIENT_OLE2.ADD_ARG(args, 1);
      CLIENT_OLE2.ADD_ARG(args, wdMove);
      CLIENT_OLE2.INVOKE(hSelection, 'MoveRight', args);
      CLIENT_OLE2.DESTROY_ARGLIST(args);*/
      
      
      
      ---- Print amt
      --AG_MESSAGE(' Print amt'||v.amount,'I');
      CLIENT_OLE2.set_property (hSelection, 'Text', v.amount); 
      args := CLIENT_OLE2.CREATE_ARGLIST;
      CLIENT_OLE2.add_arg (args, 12);
			CLIENT_OLE2.invoke (hSelection, 'MoveRight', args);
			CLIENT_OLE2.destroy_arglist (args);
   END LOOP;
   
   --------------------------
   --------------------------
   OLE2.RELEASE_OBJ(hParagraphFormat);
   OLE2.RELEASE_OBJ(hSelection);
   OLE2.RELEASE_OBJ(hDocument);
   OLE2.RELEASE_OBJ(hDocuments);
   OLE2.RELEASE_OBJ(hApplication);
   
   message('Task is Done.');
   message('Task is Done.');
EXCEPTION
   WHEN others THEN
      message('Error');
      message('Error');
END;