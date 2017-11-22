/*Lo primero es validar que efectivamente existe un bloqueo, reemplazamos el USUARIO_ORACLE por su usuario*/
select * from dba_dml_locks where owner='UXXIAC';

/*El siguiente paso es obtener la sesión que se quedo bloqueada */;
select oracle_username || ' (' || s.osuser || ')' username 
, s.sid || ',' || s.serial# sess_id 
, owner || '.' || object_name object 
, object_type 
, decode( l.block 
, 0, 'Not Blocking' 
, 1, 'Blocking' 
, 2, 'Global') status 
, decode(v.locked_mode 
, 0, 'None' 
, 1, 'Null' 
, 2, 'Row-S (SS)' 
, 3, 'Row-X (SX)' 
, 4, 'Share' 
, 5, 'S/Row-X (SSX)' 
, 6, 'Exclusive', TO_CHAR(lmode)) mode_held 
from v$locked_object v 
, dba_objects d 
, v$lock l 
, v$session s 
where v.object_id = d.object_id 
and v.object_id = l.id1 
and v.session_id = s.sid 
order by oracle_username 
, session_id; 


/*Ahora que ya se tiene la sesión se ejecuta el siguiente script para terminarla */

alter system kill session '12,10715';

alter system kill session '12,10703' IMMEDIATE;

/*Si falla terminar la sesión, debemos terminar la tarea desde una consola. 
Obtenemos el SID con la siguiente consulta: */

SELECT m.sid, m.spid, m.osuser, m.program FROM v$process p, v$session m 
WHERE m.addr = m.paddr;
/*Y ejecutamos el comando en la consola */
kill -9 58623
