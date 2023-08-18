SET OBJ = NEW IA
DIM DATA_TIME

'OBJ.SAVE_FILE  '''
OBJ.MOVING_OUT_EXCL_PRESET
OBJ.INTI
OBJ.GET_DATA
OBJ.PRCS_DATA
OBJ.PRCS_BL
OBJ.XLS
OBJ.WHI
OBJ.USG
OBJ.RW_DATA
OBJ.XLS1
OBJ.LYLOC
OBJ.T_ORD
OBJ.CHART

CreateObject("WScript.Shell").Popup "�нT�{" & DATA_TIME, 0, "�w����", &H20000
'MSGBOX "OK�F�A�нT�{" & DATA_TIME , vbExclamation + vbSystemModal

CLASS IA
	DIM MOVING_OUT_EXCL_CUST(4) '�~�h�W�h�ҥ~�q��Ȥ�}�C
    DIM FSO, CAT, CONN, SQL, RS, AR, BR, TS, BLCS, BL, CS, LOC_A, SQLL, XL, LOC_B
	DIM WS, DIR
	DIM strData
	
    FUNCTION NUM(V)
        IF ISNUMERIC(V) THEN
            NUM = CDBL(V)
          ELSE 
            NUM = 0
        END IF
    END FUNCTION
	
	Function SelectFile( )
		' File Browser via HTA
		' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
		' Features: Works in Windows Vista and up (Should also work in XP).
		'           Fairly fast.
		'           All native code/controls (No 3rd party DLL/ XP DLL).
		' Caveats:  Cannot define default starting folder.
		'           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
		'           Dialog title says "Choose file to upload".
		' Source:   https://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15ae-4ba3-bca5-ec349df65ef6/windows7-vbscript-open-file-dialog-box-fakepath?forum=ITCG

		Dim objExec, strMSHTA, wshShell

		SelectFile = ""

		' For use in HTAs as well as "plain" VBScript:
		strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
				 & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
				 & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
		' For use in "plain" VBScript only:
		' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
		'          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
		'          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

		Set wshShell = CreateObject( "WScript.Shell" )
		Set objExec = wshShell.Exec( strMSHTA )

		SelectFile = objExec.StdOut.ReadLine( )

		Set objExec = Nothing
		Set wshShell = Nothing
	End Function
	
	'�~�h�W�h�ҥ~�]�w
	'�H�U�q��Ȥᤣ����B�Ȥ�M�x��01�w
	'�̦h�i�]�w5��
	SUB MOVING_OUT_EXCL_PRESET
		'MOVING_OUT_EXCL_CUST(0) = "�K�����K"
		'MOVING_OUT_EXCL_CUST(1) = "�x�W�y��"
		MOVING_OUT_EXCL_CUST(0) = ""
		MOVING_OUT_EXCL_CUST(1) = ""
		MOVING_OUT_EXCL_CUST(2) = ""
		MOVING_OUT_EXCL_CUST(3) = ""
		MOVING_OUT_EXCL_CUST(4) = ""
	END SUB
	
	SUB SAVE_FILE
		'�}�l�e����,�קK�ϥΪ̤z�A����������O
		MsgBox("�Y�N�۰ʤU������A�L�{���Фžާ@�C")
	
		'�����ɮרt�α�����Shell�����
		set FSO = CreateObject("Scripting.FileSystemObject")
		set WS = CreateObject("WScript.Shell")
		CURR_DIR = WS.CURRENTDIRECTORY
		
		'�s�ɸ��|�����{���Ҧb��Ƨ�\�ثe���
        DIR = CURR_DIR & "\" & YEAR(DATE()) & RIGHT("0" & MONTH(DATE()), 2)  & RIGHT("0" & DAY(DATE()), 2) & "\"
		if not FSO.FolderExists(DIR) then 
            set f = FSO.CreateFolder(DIR)
        end if
		'�p���|�w��IA73R.TXT�h�R��,�קK���X�O�_�л\��ܮ�
		if FSO.FileExists(DIR & "IA73R.TXT") then fso.DeleteFile(DIR & "IA73R.TXT")

		'����IE�����
        SET IE=CREATEOBJECT("INTERNETEXPLORER.APPLICATION")
        IE.Visible=TRUE
		WS.AppActivate IE
		
		'�bURL�ѼƱa�JIA73R,�OIE�s�V^DR33
		NV = "http://eas.csc.com.tw/drw/report/drw33?reportId=IA73R"
        IE.Navigate(NV)
		
		'���ݺ������J����
        'DO WHILE IE.BUSY OR IE.READYSTATE <> 4 : LOOP
		'�������J������A�A��3��
		WSCRIPT.SLEEP 5000
		
        '���U�Ĥ@�ӤU�����s
		IE.DOCUMENT.GETELEMENTBYID("clickDownload0").CLICK

		DO WHILE IE.BUSY OR IE.READYSTATE <> 4 : LOOP
        WSCRIPT.SLEEP 5000
		WS.SENDKEYS "%n"
		WSCRIPT.SLEEP 1000
		WS.SENDKEYS "{TAB}"
		WSCRIPT.SLEEP 1000
		WS.SENDKEYS "{DOWN 2}"
		WSCRIPT.SLEEP 1000
		WS.SENDKEYS "~"
		WSCRIPT.SLEEP 1000

		'�Q�ΰŶKï�Ȧs�ɮק�����|
		DIM Form, TextBox
		set Form = CreateObject("Forms.Form.1")
		set TextBox = Form.Controls.Add("Forms.TextBox.1").Object
		TextBox.MultiLine = True
		TextBox.Text = DIR & "IA73R.TXT"
		TextBox.SelStart = 0
		TextBox.SelLength = TextBox.TextLength
		TextBox.Copy
		
        WSCRIPT.SLEEP 2000
        'Ctrl+v �K�W���x�s�ɮ׹�ܮ�
		WS.SENDKEYS "^v~"
		
        WSCRIPT.SLEEP 2000        
		'����IE����
        IE.QUIT
		set Form = NOTHING
        set IE = NOTHING
	END SUB
	
    SUB INTI
		set FSO = CreateObject("Scripting.FileSystemObject")
		set WS = CreateObject("WScript.Shell")
		'CURR_DIR = WS.CURRENTDIRECTORY
		'DIR = CURR_DIR & "\20200601\"
	
        'SET fso = CreateObject("Scripting.FileSystemObject")
		'WSCRIPT.SLEEP 2000
		'If Not FSO.FileExists(DIR & "IA73R.TXT") Then
		'	MsgBox(DIR & "IA73R.TXT" & "�ɮפ��s�b�C")
		'	WSCRIPT.QUIT
		'End If
				
        'SET F1 = FSO.OPENTEXTFILE(DIR & "IA73R.TXT")
		DIM pathFileName
		pathFileName = SelectFile()
		'MsgBox(pathFileName)
		WSCRIPT.SLEEP 1000
		IF pathFileName = "" THEN
			WSCRIPT.QUIT
		END IF
		
		'IF LEFT(pathFileName,5) <> "IA73R" THEN
		'	MsgBox("�ɦW�DIA73R�}�Y")
		'	WSCRIPT.QUIT
		'END IF
		
		IF RIGHT(pathFileName,3) <> "TXT" AND RIGHT(pathFileName,3) <> "txt" THEN
			MsgBox("�ɮ������Dtxt")
			WSCRIPT.QUIT
		END IF
		
		'SET F1 = FSO.OPENTEXTFILE(pathFileName)
		
		'read utf-8 start
		Dim objStream

		Set objStream = CreateObject("ADODB.Stream")

		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.LoadFromFile(pathFileName)

		strData = objStream.ReadText()

		objStream.Close
		
		Set objStream = Nothing
		'read utf-8 end
		
        If (fso.FileExists("IA73.accdb")) THEN fso.DeleteFile("IA73.accdb")
        SET CAT = CREATEOBJECT("ADOX.Catalog")
        CAT.Create("provider='Microsoft.ACE.OLEDB.12.0';Data Source=IA73.accdb")
        Set conn = CreateObject("ADODB.Connection")
		conn.Provider="Microsoft.ACE.OLEDB.12.0"
        conn.OPEN("IA73.accdb")
        SQL = " CREATE TABLE IA73 ( " _
            &  "OP          VARCHAR(2),  "_
            &  "�h          INT,  "_
            &  "�q��        VARCHAR(10),  "_
            &  "����        VARCHAR(6),  "_
            &  "��a        VARCHAR(4),  "_
            &  "�p          DOUBLE,  "_
            &  "�e          INT,  "_
            &  "��          INT,  "_
            &  "�Q��        VARCHAR(6),  "_
            &  "�ղ�        VARCHAR(4),  "_
            &  "��1         INT,  "_
            &  "���        VARCHAR(4),  "_
            &  "��2         INT,  "_
            &  "�~�x        VARCHAR(6),  "_
            &  "���        VARCHAR(4),  "_
            &  "�歫        INT,  "_   
            &  "LOC         VARCHAR(6),  "_
            &  "�Ȥ�        VARCHAR(10),  "_
            &  "�w�O        VARCHAR(2),  "_
            &  "�|�p        DOUBLE,  "_
            &  "�|��        INT,  "_
            &  "��歫      INT,  "_
	    &  "�i�X�f��    INT,  "_
	    &  "�q��O      VARCHAR(2),  "_
	    &  "���        VARCHAR(2),  "_
	    &  "�W����      INT,  "_
	    &  "�W�e��      INT,  "_
	    &  "����        VARCHAR(10),  "_
	    &  "����        VARCHAR(13),  "_
	    &  "�q��Ȥ�        VARCHAR(10),  "_
	    &  "�~�P�U��        VARCHAR(6),  "_
	    &  "LY����        VARCHAR(2),  "_
	    &  "�Q����        VARCHAR(2),  "_
	    &  "ORD        VARCHAR(7)  "_
            & " ) "
        CONN.EXECUTE(SQL)
        set rs = CreateObject("ADODB.recordset") 
        sql = "select * from IA73 "
        rs.open sql,conn,1,3
    END SUB
	
    SUB GET_DATA
		DIM STR
        'SS = F1.READALL
        'AR = SPLIT(SS, VBCRLF, -1, 1)
		AR = SPLIT(strData, VBCRLF, -1, 1)

        '�^����Ʈɶ��A�ín�D�T�{
		STR = ""
		FOR i = 0 TO UBOUND(MOVING_OUT_EXCL_CUST)
			iF MOVING_OUT_EXCL_CUST(i) <> "" THEN
				STR = STR & MOVING_OUT_EXCL_CUST(i) & vbCrLf
			END IF
		NEXT
		
		DATA_TIME = "��Ʈɶ��G" & MID(AR(2), 57, 20)
		
		'IF STR = "" THEN
		'	ans = MSGBOX ("��Ʈɶ��G" & MID(AR(2), 57, 20) & "  �O�_�~��?", vbYesNo)
		'ELSE
		IF STR <> "" THEN
			ans = MSGBOX ("��Ʈɶ��G" & MID(AR(2), 57, 20) & vbCrLf & vbCrLf  & "�~�h�ҥ~�q��Ȥ�:" & vbCrLf & vbCrLf & STR & vbCrLf &vbTab &vbTab & "�O�_�~��?", vbYesNo)
		END IF
		If ans = vbNo Then WSCRIPT.QUIT
	
	FOR I = 5 TO UBOUND(AR) - 1
            RS.ADDNEW
            BR = SPLIT(AR(I), ",", -1, 1)
            
            RS.FIELDS(0).VALUE = BR(0)
            
            RS.FIELDS(1).VALUE = NUM(BR(1))
            
            FOR  J = 2 TO 3
                 RS.FIELDS(J).VALUE = BR(J)
            NEXT
            
            RS.FIELDS(4).VALUE = BR(J)
            
            FOR  J = 5 TO 7
                 RS.FIELDS(J).VALUE = NUM(BR(J))
            NEXT
            FOR  J = 8 TO 9
                 RS.FIELDS(J).VALUE = BR(J)
            NEXT
            
            RS.FIELDS(10).VALUE = NUM(BR(10))
            RS.FIELDS(11).VALUE = BR(11)
            RS.FIELDS(12).VALUE = NUM(BR(12))
            
            FOR  J = 13 TO 14
                 RS.FIELDS(J).VALUE = BR(J)
            NEXT
            
            RS.FIELDS(15).VALUE = NUM(BR(15))
            

                 RS.FIELDS(16).VALUE = BR(16)
	         RS.FIELDS(17).VALUE = LEFT(BR(17),5)

              '�w�O
                 RS.FIELDS(18).VALUE = LEFT(BR(16),2)

              '�|�p
	    IF LEFT(BR(2),1) ="E" OR LEFT(BR(2),1) ="F" THEN

                IF BR(0) = " " THEN

                 RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10)) + 76

                ELSE  RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10))
                
                END IF

            ELSE  RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10))

            END IF

              '�|��
                 RS.FIELDS(20).VALUE=  NUM(BR(10)) * NUM(BR(15))

              '��歫
                 RS.FIELDS(21).VALUE=  NUM(BR(12)) * NUM(BR(15))

              '�i�X�f��
	    IF BR(3)<>"      " THEN	
		 RS.FIELDS(22).VALUE = NUM(BR(12)) * NUM(BR(15))
	    ELSEIF BR(3)="      " THEN 
		 RS.FIELDS(22).VALUE = 0
	    END IF
	    
   
              '�q��O	    
	    IF LEFT(BR(2),1) ="E" OR LEFT(BR(2),1) ="F" OR LEFT(BR(2),1) ="Q" THEN
			 RS.FIELDS(23).VALUE = "�~�P"

	       ELSEIF  LEFT(BR(2),1) ="L" OR LEFT(BR(2),1) ="D" OR LEFT(BR(2),1) ="J" THEN
	    	 RS.FIELDS(23).VALUE = "���P"

	  	   ELSEIF  LEFT(BR(2),2) ="TP" THEN 
	    	 RS.FIELDS(23).VALUE = "TP"

 	 	   ELSEIF  LEFT(BR(2),1) ="T" THEN 
	    	 RS.FIELDS(23).VALUE = "����"

	       ELSE RS.FIELDS(23).VALUE = "��L"
	 
	    END IF 
	 
            '���
	       RS.FIELDS(24).VALUE =  LEFT(BR(14),2)
             
            '�W����
	    IF NUM(BR(7)) > 13000 THEN
		RS.FIELDS(25).VALUE =  NUM(BR(10)) * NUM(BR(15))
	       ELSE
		RS.FIELDS(25).VALUE =  0
	    END IF

            '�W�e��
	    IF NUM(BR(6)) > 3000 THEN
		RS.FIELDS(26).VALUE =  NUM(BR(10)) * NUM(BR(15))
	       ELSE
		RS.FIELDS(26).VALUE =  0
	    END IF


              '����&���� 
              RS.FIELDS(27).VALUE = RIGHT(BR(18),1)

              '���� 
	      RS.FIELDS(28).VALUE = TRIM(LEFT(BR(19)  ,13))
	    
	      '���P�q��Ȥ� 
              RS.FIELDS(29).VALUE = LEFT(BR(20),10)

              '�~�P�U��Ȥ� 
	      RS.FIELDS(30).VALUE = LEFT(BR(21),6)
               
              'LY����
       
              Select Case BR(0)
              
                Case "W" 

                  IF BR(5)>28 THEN
                     RS.FIELDS(31).VALUE = "Y"
                   ELSEIF BR(5)>12.7 THEN
                     RS.FIELDS(31).VALUE = "X"
                   ELSE 
                     RS.FIELDS(31).VALUE = "W"
                  END IF
		
                Case "C" 
                  IF BR(5)>28 THEN
                     RS.FIELDS(31).VALUE = "N"
                   ELSEIF BR(5)>12.7 THEN
                     RS.FIELDS(31).VALUE = "M"
                   ELSE 
                     RS.FIELDS(31).VALUE = "L"   
                  END IF

                Case "E"
                  IF BR(5)>28 THEN
                     RS.FIELDS(31).VALUE = "V"
                   ELSEIF BR(5)>12.7 THEN
                     RS.FIELDS(31).VALUE = "U"
                   ELSE 
                     RS.FIELDS(31).VALUE = "T" 
                  END IF

                 Case "H"
                  IF BR(5)>28 THEN
                     RS.FIELDS(31).VALUE = "J"
                   ELSEIF BR(5)>12.7 THEN
                     RS.FIELDS(31).VALUE = "I"
                   ELSE 
                     RS.FIELDS(31).VALUE = "H"
                  END IF
 
                 Case "?"
                  IF BR(5)>28 THEN
                     RS.FIELDS(31).VALUE = "C"
                   ELSEIF BR(5)>12.7 THEN
                     RS.FIELDS(31).VALUE = "B"
                   ELSE 
                     RS.FIELDS(31).VALUE = "A" 
                  END IF

              End Select

            '�Q����
	       RS.FIELDS(32).VALUE =  LEFT(BR(9),2)

            'ORD�q��e7�X
	       RS.FIELDS(33).VALUE =  LEFT(BR(2),7)

            RS.MOVENEXT
        NEXT
    END SUB

    SUB PRCS_DATA
        SQL = "Select loc, ����, sum(�歫) as ���q, sum(��1) as ���� INTO �x��_�U���� from IA73 group by loc, ����" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select �w�O, sum(�|��)/1000 as ���P�`�� INTO �x��_���P�s�q from IA73 WHERE �q��O = '���P' AND �w�O  IN( '01','07','17') AND OP = ' ' group by �w�O" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select �w�O, sum(�|��)/1000 as �~�P�`�� INTO �x��_�~�P�s�q from IA73 WHERE �q��O = '�~�P' AND �w�O  IN( '01','07','17') AND OP = ' ' group by �w�O" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select �w�O, sum(�|��)/1000 as �����`�� INTO �x��_�����s�q from IA73 WHERE �q��O = '����' AND �w�O  IN( '01','07','17') AND OP = ' ' group by �w�O" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, MAX(�p) as �pMAX, MIN(�p) as �pMIN,  MAX(�e) as �eMAX, MIN(�e) as �eMIN,sum(�|��)/1000 as �`��,sum(��歫)/1000 as ��� INTO �x��_LY�w�s from IA73 WHERE �w�O  IN( '01','07','17') group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as LY�w�s�� INTO �x��_LY�� from IA73 WHERE �w�O  IN( '01','07','17') AND OP <> ' ' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as �@��LY�� INTO �x��_�@��LY from IA73 WHERE �w�O  IN( '01','07','17') AND OP = '?' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as �W���eLY�� INTO �x��_�W���eLY from IA73 WHERE �w�O  IN( '01','07','17') AND OP = 'W' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as ������LY�� INTO �x��_������LY from IA73 WHERE �w�O  IN( '01','07','17') AND OP = 'C' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as �S��LY�� INTO �x��_�S��LY from IA73 WHERE �w�O  IN( '01','07','17') AND OP = 'H' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as ���j�׭� INTO �x��_���j��LY from IA73 WHERE �w�O  IN( '01','07','17') AND OP = 'E' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE �x��_LY�w�s ADD LY�w�s DOUBLE, �@�� DOUBLE, �W���e DOUBLE, ������ DOUBLE, �S�� DOUBLE, ���j�� DOUBLE"
        CONN.EXECUTE(SQL)

        SQL = "Select �w�O, sum(�|��)/1000 as �x���`���q, sum(��歫)/1000 as ����`���q, sum(�i�X�f��)/1000 as �i�X�f���q  INTO �x��_�U�w�s�q from IA73 WHERE  �w�O  IN( '01','07','17') group by �w�O" 
        CONN.EXECUTE(SQL)
       
        'SQL = "Select DISTINCT loc, �Ȥ�, ���� INTO �x��_����M�� from IA73  group by loc,  �Ȥ�, ���� " 
        'CONN.EXECUTE(SQL)
		
		SQL = "Select loc, �Ȥ�, ����, �q��Ȥ� INTO �x��_����M�� from IA73 group by loc, �Ȥ�, ����, �q��Ȥ�"
		CONN.EXECUTE(SQL)

        SQL = "Select DISTINCT loc, �Ȥ� INTO �x��_�Ȥ�M�� from IA73  group by loc,  �Ȥ� " 
        CONN.EXECUTE(SQL)
		
		'SQL = "Select DISTINCT loc, �q��Ȥ� INTO �x��_�q��Ȥ�M�� from IA73  group by loc, �q��Ȥ�" 
        'CONN.EXECUTE(SQL)

        sql = " select loc, count(����) as ����i�� into �x��_����i�� FROM �x��_�U���� GROUP BY loc "
        CONN.EXECUTE(SQL)
   
        SQL = "Select loc, sum(�歫)/1000 as �w�}����_���q, sum(��1) as �w�}����_���� INTO �x��_�����`��  from IA73 WHERE ���� <> '      ' group by loc  " 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(�|��)/1000 as �x���`���q, sum(��1) as ����, sum(�|�p) as �x��p , MAX(�p) as �pMAX, MIN(�p) as �pMIN, sum(�i�X�f��)/1000 as �i�X�f�� , MAX(��) as ��MAX, MIN(��) as ��MIN, MAX(���) as ���MAX, MIN(���) as ���MIN INTO �x��_���q from IA73 WHERE �w�O IN('01', '07', '17', '04') group by loc " 
        CONN.EXECUTE(SQL)

        SQL = " SELECT �x��_���q.*, �x��_����i��.����i�� INTO �x�Ϥ@ FROM �x��_���q LEFT JOIN �x��_����i�� ON �x��_���q.LOC = �x��_����i��.LOC "
        CONN.EXECUTE(SQL)

        SQL = " SELECT  �x�Ϥ@.*, �x��_�����`��.�w�}����_���q, �x��_�����`��.�w�}����_���� INTO �x�ϤG FROM �x�Ϥ@ LEFT JOIN �x��_�����`�� ON �x�Ϥ@.LOC =�x��_�����`��.LOC"
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE �x�ϤG ADD �������� VARCHAR(250)"
        CONN.EXECUTE(SQL)

        'SQL = "ALTER TABLE �x��_���q ADD ������B�Ȥ� VARCHAR(250)"
		SQL = "ALTER TABLE �x��_���q ADD ������B�Ȥ� VARCHAR(250), �����q��Ȥ� VARCHAR(250)"
        CONN.EXECUTE(SQL)
		
		'SQL = "ALTER TABLE �x��_���q ADD �����q��Ȥ� VARCHAR(250)"
        'CONN.EXECUTE(SQL)

    
    END SUB

    SUB PRCS_BL
        SQL = "ALTER TABLE �x��_����M��  ADD �Ƶ� VARCHAR(11) "
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE �x��_����M��  ADD �Ƶ�1 VARCHAR(11) "
        CONN.EXECUTE(SQL)
		
		SQL = "ALTER TABLE �x��_����M��  ADD �Ƶ�2 VARCHAR(11) "
        CONN.EXECUTE(SQL)
        
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_����M�� "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
           IF LOC_A <> TS.FIELDS("LOC").VALUE THEN
                CS = MID(TS.FIELDS("�Ȥ�").VALUE, 2, 2)
				OCS = MID(TS.FIELDS("�q��Ȥ�").VALUE, 2, 4)
                TS.FIELDS("�Ƶ�").VALUE = CS & "->" & TS.FIELDS("����").VALUE & " "
	            TS.FIELDS("�Ƶ�1").VALUE = "< " &CS & " >" 
				'TS.FIELDS("�Ƶ�2").VALUE = "<" & OCS & ">"
            ELSEIF  CS  <> MID(TS.FIELDS("�Ȥ�").VALUE, 2, 2)  THEN
                CS = MID(TS.FIELDS("�Ȥ�").VALUE, 2, 2)
				OCS = MID(TS.FIELDS("�q��Ȥ�").VALUE, 2, 4)
                TS.FIELDS("�Ƶ�").VALUE = CS & "->" & TS.FIELDS("����").VALUE & " "
		        TS.FIELDS("�Ƶ�1").VALUE =  "< " &CS & " >"
				'TS.FIELDS("�Ƶ�2").VALUE = "<" & OCS & ">"
            ELSE
                TS.FIELDS("�Ƶ�").VALUE = TS.FIELDS("����").VALUE & ""
                TS.FIELDS("�Ƶ�1").VALUE = ""
				'TS.FIELDS("�Ƶ�2").VALUE = ""
            END IF
            LOC_A = TS.FIELDS("LOC").VALUE
           TS.MOVENEXT
        LOOP
		ON ERROR RESUME NEXT
		
		set Ts = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_����M�� "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
            IF LOC_A <> TS.FIELDS("LOC").VALUE THEN
				OCS = MID(TS.FIELDS("�q��Ȥ�").VALUE, 2, 4)
				TS.FIELDS("�Ƶ�2").VALUE = "<" & OCS & ">"
            ELSEIF OCS <> MID(TS.FIELDS("�q��Ȥ�").VALUE, 2, 4) THEN
                OCS = MID(TS.FIELDS("�q��Ȥ�").VALUE, 2, 4)       
		        TS.FIELDS("�Ƶ�2").VALUE = "<" & OCS & ">"
            ELSE
                TS.FIELDS("�Ƶ�2").VALUE = ""
            END IF
			'MsgBox(TS.FIELDS("�Ƶ�2").VALUE)
            LOC_A = TS.FIELDS("LOC").VALUE
            TS.MOVENEXT
        LOOP
		ON ERROR RESUME NEXT

        SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_����M�� "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
            SQL = " UPDATE �x��_���q SET ������B�Ȥ� = ������B�Ȥ� & '" & TS.FIELDS("�Ƶ�1").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
            CONN.EXECUTE(SQL)
			SQL = " UPDATE �x��_���q SET �����q��Ȥ� = �����q��Ȥ� & '" & TS.FIELDS("�Ƶ�2").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
            CONN.EXECUTE(SQL)
            TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_LY�� "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET LY�w�s =  '" & TS.FIELDS("LY�w�s��").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_�@��LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET �@�� =  '" & TS.FIELDS("�@��LY��").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_�W���eLY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET �W���e =  '" & TS.FIELDS("�W���eLY��").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_������LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET ������ =  '" & TS.FIELDS("������LY��").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_�S��LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET �S�� =  '" & TS.FIELDS("�S��LY��").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from �x��_���j��LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE �x��_LY�w�s SET ���j�� =  '" & TS.FIELDS("���j�׭�").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


    END SUB


    SUB XLS
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from �x��_���q "
         Ts.open sql,conn,1,3
         SET XL = CREATEOBJECT("EXCEL.APPLICATION")
         XL.VISIBLE = TRUE
		 WS.AppActivate  XL
         XL.WORKBOOKS.ADD
		XL.Sheets.Add

   		 XL.Activesheet.name=("�x��J�`")

         FOR I = 0 TO Ts.FIELDS.COUNT-1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts

	    XL.Columns("J:M").Select

'MSGBOX ""

    	XL.Selection.Cut
    	XL.Columns("T:T").Select
    	XL.Selection.Insert -4161

'MSGBOX ""

		XL.CELLS(1,10).VALUE="�w"
		XL.CELLS(1,11).VALUE="Column"
		XL.CELLS(1,12).VALUE="Row"
		XL.CELLS(1,13).VALUE="���A"
		XL.CELLS(1,14).VALUE="�q��O"
		XL.CELLS(1,15).VALUE="�i�X�f"
	
'MSGBOX ""

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("�x��J�`").UsedRange.Rows.Count
		XL.CELLS(K,10).VALUE="=LEFT(A" & K & ",2)"
		XL.CELLS(K,11).VALUE="=MID(A" & K & ",3,2)"
		XL.CELLS(K,12).VALUE="=RIGHT(A" & K & ",1)"
		
		' IF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXECLUDE1 &">" THEN
			' XL.CELLS(K,13).VALUE=""
		' ELSEIF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXECLUDE2 &">" THEN
			' XL.CELLS(K,13).VALUE=""
		' ELSEIF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXECLUDE3 &">" THEN
			' XL.CELLS(K,13).VALUE=""
		' ELSEIF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXECLUDE4 &">" THEN
			' XL.CELLS(K,13).VALUE=""
		' ELSEIF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXECLUDE5 &">" THEN
			' XL.CELLS(K,13).VALUE=""
		' ELSE
			' XL.CELLS(K,13).VALUE="=IF(AND(LEN(R"& K & ")>6,H"& K & "<9001),""�w�صu"",IF(AND(H" & K & "<13001,I" & K & ">9000,LEN(R" & K & ")>6),""�w�ت�"",""""))"
		' END IF
		
		XL.CELLS(K,13).VALUE="=IF(AND(LEN(R"& K & ")>6,H"& K & "<9001),""�w�صu"",IF(AND(H" & K & "<13001,I" & K & ">9000,LEN(R" & K & ")>6),""�w�ت�"",""""))"
		
		FOR i = 0 TO UBOUND(MOVING_OUT_EXCL_CUST)
			IF XL.CELLS(K,19).VALUE = "<" & MOVING_OUT_EXCL_CUST(i) &">" THEN
				XL.CELLS(K,13).VALUE=""
				Exit For
			END IF
		NEXT

        XL.CELLS(K,15).VALUE="=G" & K & "/B" & K

	NEXT
    		
      XL.Columns("O:O").Select
      XL.Selection.Style = "Percent"

'MSGBOX ""



       XL.Cells.Select	
       XL.Selection.AutoFilter
      'XL.ActiveSheet.Range("A:Q").AutoFilter 13, "<>"
      'XL.ActiveSheet.Range("A:Q").AutoFilter 10, "01"

       XL.ActiveSheet.Range("A:R").AutoFilter 10, "=07", 2, "=17"
       XL.ActiveSheet.Range("A:R").AutoFilter 12, "=A", 2, "=B"
 

'MSGBOX ""
        XL.Cells.Select	
        XL.Cells.EntireColumn.AutoFit

   XL.Columns("B:Q").Select
    'XL.Columns("B:Q").EntireColumn.AutoFit

    XL.Selection.ColumnWidth = 6.5

     END SUB
     

SUB XLS1
         XL.Sheets("�x��J�`").Select

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("�x��J�`").UsedRange.Rows.Count
		
		XL.CELLS(K,14).VALUE="=IF(COUNTIFS(IA73!Q:Q,�x��J�`!A" & K & ",IA73!X:X,""�~�P"")>0,""�~�P"","""")"
		
	NEXT

    XL.Sheets("�U�w�s�q").Select

   	XL.CELLS(1,5).VALUE = "���P�`��"
	XL.CELLS(2,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""���P"",IA73!A:A,"" "")/1000"
	XL.CELLS(3,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""���P"",IA73!A:A,"" "")/1000"
	XL.CELLS(4,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""���P"",IA73!A:A,"" "")/1000"

   	XL.CELLS(1,6).VALUE = "�~�P�`��"
	XL.CELLS(2,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""�~�P"",IA73!A:A,"" "")/1000"
	XL.CELLS(3,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""�~�P"",IA73!A:A,"" "")/1000"
	XL.CELLS(4,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""�~�P"",IA73!A:A,"" "")/1000"

   	XL.CELLS(1,7).VALUE = "���P�i�X"
	XL.CELLS(2,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""01"",IA73!A:A,"" "",IA73!X:X,""���P"")/1000"
	XL.CELLS(3,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""07"",IA73!A:A,"" "",IA73!X:X,""���P"")/1000"
	XL.CELLS(4,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""17"",IA73!A:A,"" "",IA73!X:X,""���P"")/1000"

   	XL.CELLS(1,8).VALUE = "�~�P�i�X"
	XL.CELLS(2,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""01"",IA73!A:A,"" "",IA73!X:X,""�~�P"")/1000"
	XL.CELLS(3,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""07"",IA73!A:A,"" "",IA73!X:X,""�~�P"")/1000"
	XL.CELLS(4,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""17"",IA73!A:A,"" "",IA73!X:X,""�~�P"")/1000"
	
   	XL.CELLS(1,10).VALUE = "���q�`��"
	XL.CELLS(2,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""���P"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(3,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""���P"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(4,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""���P"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(5,10).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
	XL.CELLS(6,10).FormulaR1C1 = "=R[-1]C/R5C2"
	XL.CELLS(6,10).Style = "Percent"

   	XL.CELLS(1,11).VALUE = "�w�ت�"
	XL.CELLS(2,11).VALUE = "=SUMIFS(�x��J�`!B:B,�x��J�`!M:M,""�w�ت�"",�x��J�`!J:J,""01"",�x��J�`!N:N,"""")"

	XL.CELLS(1,12).VALUE = "�w�صu"
	XL.CELLS(2,12).VALUE = "=SUMIFS(�x��J�`!B:B,�x��J�`!M:M,""�w�صu"",�x��J�`!J:J,""01"",�x��J�`!N:N,"""")"

	XL.CELLS(3,13).VALUE = "�~�P17"
	XL.CELLS(4,13).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""�~�P"",IA73!A:A,"" "",IA73!H:H,""<9001"",IA73!C:C,""<>QJ*"")/1000"

	XL.CELLS(1,13).VALUE = "�~�P07"
	XL.CELLS(2,13).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""�~�P"",IA73!A:A,"" "",IA73!H:H,""<13000"")/1000-M4"

	XL.CELLS(3,14).VALUE = "=""�U�w�w�s�G01�w�G""& ROUND(B2,0) & ""���B07�w�G"" & ROUND(B3,0) & ""���B17�w�G"" & ROUND(B4,0) & ""���A�`�w�s�G"" & ROUND(B5,0) & ""���C"""
	'XL.CELLS(4,14).VALUE = "=""   �ݥ~�h�q�G07�w(���P"" & ROUND(K2,0) & ""��+�~�P"" & ROUND(M2,0) & ""��)�B17�w("" & ROUND(L2,0) & ""��)�C"""
	XL.CELLS(4,14).VALUE = "=""   �ݥ~�h�q�G07�w(���P"" & ROUND(K2,0) & ""��+�~�P"" & ROUND(M2,0) & ""��)�B17�w(���P"" & ROUND(L2,0) & ""��+�~�P"" & ROUND(M4,0) & ""��)�C"""
	'XL.CELLS(5,14).VALUE = "=""   ���q�w�s�G01�w�G"" & ROUND(J2,0) & ""���B07�w�G"" & ROUND(J3,0) & ""���B17�w�G"" & ROUND(J4,0) & ""���C"""
	
	STR = ""
	IF XL.CELLS(2,10).VALUE > 0 THEN
	    STR = "=""   ���q�w�s�G01�w�G"" & ROUND(J2,0) & ""��"
	    IF XL.CELLS(3,10).VALUE > 0 THEN
		    STR = STR & "�B07�w�G"" & ROUND(J3,0) & ""��"
		    IF XL.CELLS(4,10).VALUE > 0 THEN
			    STR = STR & "�B17�w�G"" & ROUND(J4,0) & ""��"
			END IF
		ELSEIF XL.CELLS(4,10).VALUE > 0 THEN
		    STR = STR & """�B17�w�G"" & ROUND(J4,0) & ""��"
		END IF
	ELSEIF XL.CELLS(3,10).VALUE > 0 THEN
	    STR = "=""   ���q�w�s�G07�w�G"" & ROUND(J3,0) & ""��"
		IF XL.CELLS(7,10).VALUE > 0 THEN
		    STR = STR & "�B17�w�G"" & ROUND(J4,0) & ""��"
		END IF
	ELSEIF  XL.CELLS(4,10).VALUE > 0 THEN
	    STR = "=""   ���q�w�s�G17�w�G"" & ROUND(J4,0) & ""��"
	END IF
	
	IF XL.CELLS(2,10).VALUE > 0 OR XL.CELLS(3,10).VALUE > 0 OR XL.CELLS(4,10).VALUE > 0 THEN
		XL.CELLS(5,14).VALUE = STR & "�C"""
	END IF
	
    XL.Cells.Select	
    XL.Cells.EntireColumn.AutoFit

END SUB    

    SUB RW_DATA

    ON ERROR RESUME NEXT

         set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from IA73 "
         Ts.open sql,conn,1,3
		XL.Sheets.Add

   		XL.Activesheet.name=("IA73")
   		
   		FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts
                  
         SET Ts = NOTHING

        XL.Cells.Select	
        XL.Cells.EntireColumn.AutoFit

 

    XL.Selection.AutoFilter

    XL.ActiveSheet.Range("$A$1:$AC$11422").AutoFilter  29,"=APPLY*", 1

    XL.ActiveSheet.Range("$A$1:$AC$11422").AutoFilter  1,  "<>"
         
    XL.Sheets("�U�w�s�q").Select

'�U�����w�s
        'XL.Cells(10, 1).Select
        
        XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "�U�w�s�q!R8C1", "�ϯä��R��10", 3
      
    With XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotFields("���")
        .Orientation = 1
        .Position = 1
    End With
    With XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotFields("�w�O")
        .Orientation = 2
        .Position = 1
    End With

   XL.ActiveSheet.PivotTables("�ϯä��R��10").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotFields("�|��"), "�[�` - �|��",  -4157

   XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotFields("�[�` - �|��").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("�ϯä��R��10").CompactLayoutColumnHeader = "�w�O"
    XL.ActiveSheet.PivotTables("�ϯä��R��10").CompactLayoutRowHeader = "���"
    XL.ActiveSheet.PivotTables("�ϯä��R��10").DataPivotField.PivotItems("�[�` - �|��").Caption = "�w�s��"

    With XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotFields("�w�O")
ON ERROR RESUME NEXT
        .PivotItems("02").Visible = False
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

 XL.ActiveSheet.PivotTables("�ϯä��R��10").PivotSelect "", 0, True
    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With

'�U�Q����w�s

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "�U�w�s�q!R23C1", "�ϯä��R��13", 3
   

    With XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotFields("�w�O")
        .Orientation = 2
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotFields("�Q����")
        .Orientation = 1
        .Position = 1
    End With

    XL.ActiveSheet.PivotTables("�ϯä��R��13").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotFields("�|��"), "�[�` - �|��",  -4157

    XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotFields("�[�` - �|��").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("�ϯä��R��13").CompactLayoutColumnHeader = "�w�O"


    XL.ActiveSheet.PivotTables("�ϯä��R��13").CompactLayoutRowHeader = "�Q����"
    XL.ActiveSheet.PivotTables("�ϯä��R��13").DataPivotField.PivotItems("�[�` - �|��").Caption = "�Q����"


   With XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotFields("�w�O")
   ON ERROR RESUME NEXT
        .PivotItems("02").Visible = False
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("(blank)").Visible = False
   End With

 XL.ActiveSheet.PivotTables("�ϯä��R��13").PivotSelect "", 0, True
    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With

'�����q��U�����w�s
   XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "�U�w�s�q!R29C8", "�ϯä��R��11", 3
      
    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("���")
        .Orientation = 1
        .Position = 1
    End With
    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�w�O")
        .Orientation = 2
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�q��O")
        .Orientation = 3
        .Position = 1
    End With

  With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("OP")
        .Orientation = 3
        .Position = 2
    End With
 
    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�q��O")
        .PivotItems("TP").Visible = False
        .PivotItems("���P").Visible = False
        .PivotItems("�~�P").Visible = False
        .PivotItems("��L").Visible = False
        .PivotItems("(blank)").Visible = False

    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("OP")
ON ERROR RESUME NEXT
        .PivotItems("?").Visible = False
        .PivotItems("C").Visible = False
        .PivotItems("H").Visible = False
        .PivotItems("R").Visible = False
        .PivotItems("S").Visible = False
        .PivotItems("W").Visible = False
        .PivotItems("E").Visible = False
    End With

   XL.ActiveSheet.PivotTables("�ϯä��R��11").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�|��"), "�[�` - �|��",  -4157

   XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�[�` - �|��").NumberFormat = "#,##0,"

   XL.ActiveSheet.PivotTables("�ϯä��R��11").CompactLayoutColumnHeader = "�w�O"
    XL.ActiveSheet.PivotTables("�ϯä��R��11").CompactLayoutRowHeader = "���"
    XL.ActiveSheet.PivotTables("�ϯä��R��11").DataPivotField.PivotItems("�[�` - �|��").Caption = "���⭫"
   
    With XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotFields("�w�O")
        .PivotItems("02").Visible = False '550
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

    XL.ActiveSheet.PivotTables("�ϯä��R��11").PivotSelect "", 0, True
    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With

    XL.Range("N7").Select
    XL.ActiveSheet.Hyperlinks.Add XL.Selection, "http://iscm.csc.com.tw/erp/se/do?_pageId=sejjTU15"


    'LEEWAY������

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C32", 3).CreatePivotTable "�U�w�s�q!R8C8", "�ϯä��R��12", 3
    'XL.Sheets("�U�w�s�q").Select


    With XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("LY����")
        .Orientation = 1
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("�w�O")
        .Orientation = 2
        .Position = 1
    End With

    XL.ActiveSheet.PivotTables("�ϯä��R��12").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("�|��"), "�[�` - �|��", -4157

    XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("�[�` - �|��").NumberFormat = "#,##0,"

    With XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("�w�O")

    XL.ActiveSheet.PivotTables("�ϯä��R��12").DataPivotField.PivotItems("�[�` - �|��").Caption = "LY��"

ON ERROR RESUME NEXT
        .PivotItems("02").Visible = False
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotFields("LY����")
        .PivotItems("(blank)").Visible = False
    End With


    XL.ActiveSheet.PivotTables("�ϯä��R��12").PivotSelect "", 0, True
    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = False
    End With


	XL.Cells.Select	  
        XL.Cells.EntireColumn.AutoFit
         
    END SUB

	 SUB USG
		XL.Sheets.Add

   		XL.Activesheet.name=("�x��ϥ�")

		XL.ActiveWorkbook.PivotCaches.Create (1,"�x��J�`!R1C1:R1000C15", 3).CreatePivotTable "�x��ϥ�!R3C1", "�ϯä��R��4", 3
  
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").Orientation = 3
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").Position = 1

		XL.ActiveSheet.PivotTables("�ϯä��R��4").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�x���`���q"), "�p�� - ����",  -4112

		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("Row").Orientation =1
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("Row").Position = 1

		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("Column").Orientation = 2
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("Column").Position = 1
 
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").PivotItems("07").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").PivotItems("17").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").PivotItems("(blank)").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").PivotItems("04").Visible = False
 
		XL.ActiveSheet.PivotTables("�ϯä��R��4").PivotFields("�w").EnableMultiplePageItems = True

		' XL.ActiveWorkbook.PivotCaches.Create (1,"�x��J�`!R1C1:R1000C15", 3).CreatePivotTable "�x��ϥ�!R16C1", "�ϯä��R��5", 3
  
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").Orientation = 3
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").Position = 1

		' XL.ActiveSheet.PivotTables("�ϯä��R��5").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�x���`���q"), "�p�� - ����",  -4112

		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("Row").Orientation =1
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("Row").Position = 1

		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("Column").Orientation = 2
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("Column").Position = 1
 
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").PivotItems("01").Visible = False
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").PivotItems("17").Visible = False
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").PivotItems("(blank)").Visible = False
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").PivotItems("04").Visible = False
 
		' XL.ActiveSheet.PivotTables("�ϯä��R��5").PivotFields("�w").EnableMultiplePageItems = True

		XL.ActiveWorkbook.PivotCaches.Create (1,"�x��J�`!R1C1:R1000C15", 3).CreatePivotTable "�x��ϥ�!R32C1", "�ϯä��R��6", 3
  
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").Orientation = 3
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").Position = 1

		XL.ActiveSheet.PivotTables("�ϯä��R��6").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�x���`���q"), "�p�� - ����",  -4112

		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("Row").Orientation =1
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("Row").Position = 1

		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("Column").Orientation = 2
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("Column").Position = 1
 
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").PivotItems("01").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").PivotItems("07").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").PivotItems("(blank)").Visible = False
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").PivotItems("04").Visible = False
 
		XL.ActiveSheet.PivotTables("�ϯä��R��6").PivotFields("�w").EnableMultiplePageItems = True
		XL.Cells.Select
		XL.Selection.ColumnWidth = 2.50


     END SUB
    
    SUB WHI
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from �x��_�U�w�s�q "
         Ts.open sql,conn,1,3

		XL.Sheets.Add

   		XL.Activesheet.name=("�U�w�s�q")
   		
   		FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts
                  
         SET Ts = NOTHING
         
         
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select ���P�`�� from �x��_���P�s�q "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+7).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         
         XL.CELLS(2,7).COPYFROMRECORDSET Ts
         SET Ts = NOTHING
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select �~�P�`�� from �x��_�~�P�s�q "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+8).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         XL.CELLS(2,8).COPYFROMRECORDSET Ts
         SET Ts = NOTHING
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select �����`�� from �x��_�����s�q "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+9).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         
         XL.CELLS(2,9).COPYFROMRECORDSET Ts
  



        XL.Cells.Select	
        XL.Selection.NumberFormatLocal = "#,##0_ "
        XL.Cells.EntireColumn.AutoFit
        
           FOR J = 2 TO XL.ActiveWorkbook.Worksheets("�U�w�s�q").UsedRange.Columns.Count       
           XL.CELLS(5,J).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
		   XL.CELLS(6,J).FormulaR1C1 = "=R[-1]C/R5C2"
		   XL.CELLS(6,J).Style = "Percent"
     
         NEXT
         
   	END SUB


SUB CHART
   	WITH XL
   		.Sheets("�x��J�`").Select
		.Range("A:A,D:D").Select
		.Charts.Add  '�x��p
                .Sheets("Chart1").Name = "�x��p"
   	END WITH
   	
END SUB


SUB LYLOC

        set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from �x��_LY�w�s "
         Ts.open sql,conn,1,3

		XL.Sheets.Add

   		XL.Activesheet.name=("LY�x��")
   		
   	 FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts
                  
         SET Ts = NOTHING

		XL.CELLS(1,14).VALUE  ="����"
		XL.CELLS(1,15).VALUE = "LY��"

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("LY�x��").UsedRange.Rows.Count
		XL.CELLS(K,14).VALUE="=G" & K & "/F" & K
		XL.CELLS(K,15).VALUE="=H" & K & "/F" & K
	NEXT

		XL.CELLS(1,16).VALUE  ="�p��"
		XL.CELLS(1,17).VALUE = "�e��"
		XL.CELLS(1,18).VALUE = "����"
		XL.CELLS(1,19).VALUE = "AP�O"
		XL.CELLS(1,20).VALUE = "PX�O"


	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("LY�x��").UsedRange.Rows.Count

                XL.CELLS(K,16).VALUE="=IF(B" & K & "<=12.7,"""",IF(AND(B" & K & "<=28,C" & K & ">12.7),"""",IF(C" & K & ">28,"""",IF(LEFT(A" & K & ",2)=""01"","""",""�p���`""))))"

                XL.CELLS(K,17).VALUE="=IF(LEFT(A" & K & ",2)<>""01"",IF(D" & K & ">3250,""�e���`"",""""),"""")"

                XL.CELLS(K,18).VALUE="=IF(AND(COUNTBLANK(I" & K & ":M" & K & ")<4,LEFT(A" & K & ",2)<>""01""),""�V�x"","""")"

                XL.CELLS(K,19).VALUE="=IF(COUNTIFS(IA73!Q:Q,LY�x��!A" & K & ",IA73!AC:AC,""APPLY HEAT"")>0,""AP�O *""& COUNTIFS(IA73!Q:Q,LY�x��!A" & K & ",IA73!AC:AC,""APPLY HEAT""),"""")"            

                XL.CELLS(K,20).VALUE="=IF(COUNTIFS(IA73!Q:Q,LY�x��!A" & K & ",IA73!AC:AC,""PX1"")>0,""PX�O*""& COUNTIFS(IA73!Q:Q,LY�x��!A" & K & ",IA73!AC:AC,""PX1""),"""")"            
	NEXT
	
  	NN=XL.ActiveWorkbook.Worksheets("LY�x��").UsedRange.Rows.Count

       XL.Range("A1:T" & NN ).Select

    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With


    With XL.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = False
    End With

    XL.Selection.AutoFilter
    XL.ActiveSheet.Range("A:O").AutoFilter  15,  ">=0.6", 1

      XL.Columns("N:O").Select
      XL.Selection.Style = "Percent"

      XL.Cells.Select	
      XL.Selection.ColumnWidth = 7
   

END SUB

SUB T_ORD

		XL.Sheets.Add

   		XL.Activesheet.name=("�����q��")

    '�U��줺���q��M��

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C34", 3).CreatePivotTable "�����q��!R6C1", "�ϯä��R��14", 3
   
    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("ORD")
        .Orientation = 1
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�Ȥ�")
        .Orientation = 2
        .Position = 1
    End With


    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("OP")
        .Orientation = 3
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�q��O")
        .Orientation = 3
        .Position = 2
    End With

    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�w�O")
        .Orientation = 3
        .Position = 3
    End With




 
    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�q��O")
	ON ERROR RESUME NEXT
        .PivotItems("TP").Visible = False
        .PivotItems("���P").Visible = False
        .PivotItems("�~�P").Visible = False
        .PivotItems("��L").Visible = False
        .PivotItems("(blank)").Visible = False

    End With

    XL.ActiveSheet.PivotTables("�ϯä��R��14").AddDataField XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�|��"), "�[�` - �|��", -4157

    XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�[�` - �|��").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("�ϯä��R��14").DataPivotField.PivotItems("�[�` - �|��").Caption = "������"

    With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("�w�O")

    ON ERROR RESUME NEXT
        .PivotItems("02").Visible = False
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

 With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("OP")
ON ERROR RESUME NEXT
        .PivotItems("?").Visible = False
        .PivotItems("C").Visible = False
        .PivotItems("H").Visible = False
        .PivotItems("R").Visible = False
        .PivotItems("S").Visible = False
        .PivotItems("W").Visible = False
        .PivotItems("E").Visible = False
    End With


  '  With XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotFields("LY����")
  '      .PivotItems("(blank)").Visible = False
  '  End With


    XL.ActiveSheet.PivotTables("�ϯä��R��14").PivotSelect "", 0, True
    XL.Selection.Borders(5).LineStyle = -4142
    XL.Selection.Borders(6).LineStyle = -4142
    With XL.Selection.Borders(7)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(8)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(10)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(11)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection.Borders(12)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    With XL.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = False
    End With


	XL.Cells.Select	  
        XL.Cells.EntireColumn.AutoFit

 XL.Sheets("�U�w�s�q").Select
END SUB



END CLASS