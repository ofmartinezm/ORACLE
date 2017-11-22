http://formsmigration.com/webutil1.html

Webutil Example Code Fragments
Most of these code fragments are from our own code. One or two are from publicly available web sites, a couple are from the webutil example forms provided by Oracle Corp. 
Open file for writing
filename := 'c:\temp\'||to_char(in_load_sequence)||'.CTL';
MYFILE := client_text_io.FOPEN(FILENAME, 'W');

Write to the file
--disp_message('writing file '||filename);
client_text_io.putf(myfile, 'load DATA append INTO TABLE SPOT_LOAD1');
client_text_io.putf(myfile, ' FIELDS TERMINATED BY '||''''||','||''''||
' optionally enclosed by '||''''||'"'||''''||' TRAILING NULLCOLS');
client_text_io.putf(myfile, ' ( ');
client_text_io.putf(myfile, 'COL1 CHAR NULLIF (COL1=BLANKS)');
client_text_io.putf(myfile, ',COL2 CHAR NULLIF (COL2=BLANKS)');
client_text_io.putf(myfile, ',COL3 CHAR NULLIF (COL3=BLANKS)');
client_text_io.putf(myfile, ',COL4 CHAR NULLIF (COL4=BLANKS)');
client_text_io.putf(myfile, ',COL5 CHAR NULLIF (COL5=BLANKS)');
client_text_io.putf(myfile, ',COL6 CHAR NULLIF (COL6=BLANKS)');
client_text_io.putf(myfile, ',COL7 CHAR NULLIF (COL7=BLANKS)');
client_text_io.putf(myfile, ',COL8 CHAR NULLIF (COL8=BLANKS)');
client_text_io.putf(myfile, ',COL9 CHAR NULLIF (COL9=BLANKS)');
client_text_io.putf(myfile, ',COL10 CHAR NULLIF (COL10=BLANKS)');
client_text_io.putf(myfile, ', COUNTRY CONSTANT '||''''||:b2.country||''''||' ');
client_text_io.putf(myfile, ', LOAD_SEQUENCE CONSTANT '||to_char(in_load_sequence));
client_text_io.putf(myfile, ', RAW_FILE_LINE sequence (max,1)');
client_text_io.putf(myfile, ', LOAD_USER CONSTANT '||''''||user||''''|| ' ');
client_text_io.putf(myfile, ', LOAD_DATE CONSTANT '||''''||
to_char(sysdate,'DD-MON-YYYY')||''''||' ');
client_text_io.putf(myfile, ', FMT6_FILE_NUMBER CONSTANT 1 )');

Close the file
/* OK close the file and copy to the right place */
client_text_io.fclose(myfile); 
--disp_message('closing file '||filename);
synchronize;
--disp_message('Check status of c:\temp .ctl file'); 
--status := webutil_file_transfer.as_to_client('c:\temp\'||to_char(in_load_sequence), 'c:\temp\'||to_char(in_load_sequence));
client_host('cmd.exe /c copy '||'c:\temp\'||to_char(in_load_sequence)||'.CTL'||' '||'t:\dataload');

Client copy file
declare 
file_done boolean;

file_done := client_copy_file(:b2.FILE_NAME , OUT_FILE || '6_1.dat');

Client Host
client_host('NOTEPAD.EXE '||'I:\INPUT\ADEX_INT\'||TO_CHAR(:b2.LOAD_SEQUENCE)||'_'||V_RAW_TABLE||'.dis',0);

Client Sample 2

PROCEDURE convert_to_mp3(v_country varchar2, v_file_name varchar2) IS
--
-- convert creative file from .WAV format to MP3 format
-- Files are assumed to be in a directory structure on the R: drive
-- R:\\
host_string varchar2(300);
out_file client_Text_IO.File_Type;
linebuf VARCHAR2(1800);
filename VARCHAR2(30);
BEGIN
filename:= 'c:\temp\mp3conv'||to_char(:b1.file_countx)||'.cmd';
host_string := 'cmd.exe /c '||'c:\mp3enc\mp3enc.exe r:\'||v_country||'\'||v_file_name||' r:\'||v_country||'\'|| ' >c:\temp\mp3.out' ;
client_host(host_string); 
:b1.file_countx := :b1.file_countx + 1;
END;

Checking file existance

declare
f_exists boolean; 
begin
temp_file := 'r:\'||c1rec.country||'\'||prefix||chr(96+i)||'.'||suffix;
f_exists := webutil_file.file_exists(temp_file);

if f_exists then
update ......

end if;


File Selection Dialogue
default_value('l:\','global.dir');
:b2.file_name := webutil_file.file_selection_dialog( :global.dir, null ,null,'Find File');
if instr(:b2.file_name,' ') > 0 then
err_message('File path or file name as a space character within it');
end if;
:global.dir := :b2.file_name;


Download from the App Server to the Client
PROCEDURE DOWNLOAD_AS IS
l_success boolean;
l_bare_filename varchar2(50);
BEGIN
--l_bare_filename := substr(:upload.file_name,instr(:download.file_name,'\',-1)+1);
l_success := webutil_file_transfer.AS_to_Client_with_progress
(clientFile => :download.file_name
,serverFile => 'd:\temp\downloaded_from_as.txt'
,progressTitle => 'Download from Application Server in progress'
,progressSubTitle => 'Please wait'
);
if l_success
then
message('File downloaded successfully from the Application Server');
else
message('File download from Application Server failed');
end if;

exception
when others
then
message('File download failed: '||sqlerrm);

END;

Download file from DB platform to the Client
PROCEDURE DOWNLOAD_DB IS
l_success boolean;
BEGIN
l_success := webutil_file_transfer.DB_To_Client_with_progress
(clientFile => :download.file_name
,tableName => 'WU_TEST_TABLE'
,columnName => 'BLOB'
,whereClause => 'ID = 1'
,progressTitle => 'Download from Database in progress'
,progressSubTitle=> 'Please wait'
);
if l_success
then
message('File downloaded successfully from the Database');
else
message('File download from Database failed');
end if;

exception
when others
then
message('File download failed: '||sqlerrm);

END;

Get Client Information
PROCEDURE GET_CLIENTINFO IS
BEGIN
:CLIENTINFO.USER_NAME := webutil_clientinfo.get_user_name;
:CLIENTINFO.IP_ADDRESS := webutil_clientinfo.get_ip_address;
:CLIENTINFO.HOST_NAME := webutil_clientinfo.get_host_name;
:CLIENTINFO.OPERATING_SYSTEM := webutil_clientinfo.get_operating_system;
:CLIENTINFO.JAVA_VERSION := webutil_clientinfo.get_java_version;
:CLIENTINFO.PATH_SEPERATOR := webutil_clientinfo.get_path_separator;
:CLIENTINFO.FILE_SEPERATOR := webutil_clientinfo.get_file_separator;
:CLIENTINFO.LANGUAGE := webutil_clientinfo.get_language;
:CLIENTINFO.TIME_ZONE := webutil_clientinfo.get_time_zone;
:CLIENTINFO.DATE_TIME := webutil_clientinfo.get_date_time;
END;

OLE Writing
PROCEDURE OLE_WRITE IS
app CLIENT_OLE2.OBJ_TYPE;
docs CLIENT_OLE2.OBJ_TYPE; 
doc CLIENT_OLE2.OBJ_TYPE; 
selection CLIENT_OLE2.OBJ_TYPE; 
args CLIENT_OLE2.LIST_TYPE;
BEGIN
-- create a new document
app := CLIENT_OLE2.CREATE_OBJ('Word.Application');
if :ole.silent = 'Y' 
then
CLIENT_OLE2.SET_PROPERTY(app,'Visible',0);
else
CLIENT_OLE2.SET_PROPERTY(app,'Visible',1);
end if;
docs := CLIENT_OLE2.GET_OBJ_PROPERTY(app, 'Documents');
doc := CLIENT_OLE2.INVOKE_OBJ(docs, 'add');

selection := CLIENT_OLE2.GET_OBJ_PROPERTY(app, 'Selection');

-- insert data into new document from long item
CLIENT_OLE2.SET_PROPERTY(selection, 'Text', :ole.oletext);

-- save document as example.tmp
args := CLIENT_OLE2.CREATE_ARGLIST;
CLIENT_OLE2.ADD_ARG(args, :ole.filename);
CLIENT_OLE2.INVOKE(doc, 'SaveAs', args);
CLIENT_OLE2.DESTROY_ARGLIST(args);

-- close example.tmp
args := CLIENT_OLE2.CREATE_ARGLIST;
CLIENT_OLE2.ADD_ARG(args, 0);
CLIENT_OLE2.INVOKE(doc, 'Close', args);
CLIENT_OLE2.DESTROY_ARGLIST(args);

CLIENT_OLE2.RELEASE_OBJ(selection);
CLIENT_OLE2.RELEASE_OBJ(doc); 
CLIENT_OLE2.RELEASE_OBJ(docs); 

-- exit MSWord 
CLIENT_OLE2.INVOKE(app,'Quit');

END;

Other Stuff
PROCEDURE separate_frame_options IS
BEGIN
if WebUtil_SeparateFrame.IsSeparateFrame
then
WebUtil_SeparateFrame.SetTitle('WebUtil Demo Form');
end if;
END;

Upload file from Client Platform to the Appserver Platform
PROCEDURE UPLOAD_AS IS
l_success boolean;
l_bare_filename varchar2(50);
BEGIN
l_bare_filename := substr(:upload.file_name,instr(:upload.file_name,'\',-1)+1);
l_success := webutil_file_transfer.Client_To_AS_with_progress
(clientFile => :upload.file_name
,serverFile => 'd:\temp\'||l_bare_filename
,progressTitle => 'Upload to Application Server in progress'
,progressSubTitle => 'Please wait'
,asynchronous => false
,callbackTrigger => null
);
if l_success
then
message('File uploaded successfully to the Application Server');
else
message('File upload to Application Server failed');
end if;

exception
when others
then
message('File upload failed: '||sqlerrm);
END;
