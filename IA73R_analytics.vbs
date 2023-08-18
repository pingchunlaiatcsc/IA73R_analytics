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

CreateObject("WScript.Shell").Popup "請確認" & DATA_TIME, 0, "已完成", &H20000
'MSGBOX "OK了，請確認" & DATA_TIME , vbExclamation + vbSystemModal

CLASS IA
	DIM MOVING_OUT_EXCL_CUST(4) '外搬規則例外訂單客戶陣列
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
	
	'外搬規則例外設定
	'以下訂單客戶不分交運客戶專儲於01庫
	'最多可設定5組
	SUB MOVING_OUT_EXCL_PRESET
		'MOVING_OUT_EXCL_CUST(0) = "春源鋼鐵"
		'MOVING_OUT_EXCL_CUST(1) = "台灣造船"
		MOVING_OUT_EXCL_CUST(0) = ""
		MOVING_OUT_EXCL_CUST(1) = ""
		MOVING_OUT_EXCL_CUST(2) = ""
		MOVING_OUT_EXCL_CUST(3) = ""
		MOVING_OUT_EXCL_CUST(4) = ""
	END SUB
	
	SUB SAVE_FILE
		'開始前提示,避免使用者干涉按鍵模擬指令
		MsgBox("即將自動下載報表，過程中請勿操作。")
	
		'產生檔案系統控制物件及Shell控制物件
		set FSO = CreateObject("Scripting.FileSystemObject")
		set WS = CreateObject("WScript.Shell")
		CURR_DIR = WS.CURRENTDIRECTORY
		
		'存檔路徑為此程式所在資料夾\目前日期
        DIR = CURR_DIR & "\" & YEAR(DATE()) & RIGHT("0" & MONTH(DATE()), 2)  & RIGHT("0" & DAY(DATE()), 2) & "\"
		if not FSO.FolderExists(DIR) then 
            set f = FSO.CreateFolder(DIR)
        end if
		'如路徑已有IA73R.TXT則刪除,避免跳出是否覆蓋對話框
		if FSO.FileExists(DIR & "IA73R.TXT") then fso.DeleteFile(DIR & "IA73R.TXT")

		'產生IE控制物件
        SET IE=CREATEOBJECT("INTERNETEXPLORER.APPLICATION")
        IE.Visible=TRUE
		WS.AppActivate IE
		
		'在URL參數帶入IA73R,令IE連向^DR33
		NV = "http://eas.csc.com.tw/drw/report/drw33?reportId=IA73R"
        IE.Navigate(NV)
		
		'等待網頁載入完成
        'DO WHILE IE.BUSY OR IE.READYSTATE <> 4 : LOOP
		'網頁載入完成後，再等3秒
		WSCRIPT.SLEEP 5000
		
        '按下第一個下載按鈕
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

		'利用剪貼簿暫存檔案完整路徑
		DIM Form, TextBox
		set Form = CreateObject("Forms.Form.1")
		set TextBox = Form.Controls.Add("Forms.TextBox.1").Object
		TextBox.MultiLine = True
		TextBox.Text = DIR & "IA73R.TXT"
		TextBox.SelStart = 0
		TextBox.SelLength = TextBox.TextLength
		TextBox.Copy
		
        WSCRIPT.SLEEP 2000
        'Ctrl+v 貼上至儲存檔案對話框
		WS.SENDKEYS "^v~"
		
        WSCRIPT.SLEEP 2000        
		'關閉IE視窗
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
		'	MsgBox(DIR & "IA73R.TXT" & "檔案不存在。")
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
		'	MsgBox("檔名非IA73R開頭")
		'	WSCRIPT.QUIT
		'END IF
		
		IF RIGHT(pathFileName,3) <> "TXT" AND RIGHT(pathFileName,3) <> "txt" THEN
			MsgBox("檔案類型非txt")
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
            &  "層          INT,  "_
            &  "訂單        VARCHAR(10),  "_
            &  "提單        VARCHAR(6),  "_
            &  "交地        VARCHAR(4),  "_
            &  "厚          DOUBLE,  "_
            &  "寬          INT,  "_
            &  "長          INT,  "_
            &  "吊移        VARCHAR(6),  "_
            &  "調移        VARCHAR(4),  "_
            &  "片1         INT,  "_
            &  "放行        VARCHAR(4),  "_
            &  "片2         INT,  "_
            &  "外儲        VARCHAR(6),  "_
            &  "交期        VARCHAR(4),  "_
            &  "單重        INT,  "_   
            &  "LOC         VARCHAR(6),  "_
            &  "客戶        VARCHAR(10),  "_
            &  "庫別        VARCHAR(2),  "_
            &  "疊厚        DOUBLE,  "_
            &  "疊重        INT,  "_
            &  "放行重      INT,  "_
	    &  "可出貨重    INT,  "_
	    &  "訂單別      VARCHAR(2),  "_
	    &  "交月        VARCHAR(2),  "_
	    &  "超長重      INT,  "_
	    &  "超寬重      INT,  "_
	    &  "側邊        VARCHAR(10),  "_
	    &  "鋼種        VARCHAR(13),  "_
	    &  "訂單客戶        VARCHAR(10),  "_
	    &  "外銷下游        VARCHAR(6),  "_
	    &  "LY分類        VARCHAR(2),  "_
	    &  "吊移月        VARCHAR(2),  "_
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

        '擷取資料時間，並要求確認
		STR = ""
		FOR i = 0 TO UBOUND(MOVING_OUT_EXCL_CUST)
			iF MOVING_OUT_EXCL_CUST(i) <> "" THEN
				STR = STR & MOVING_OUT_EXCL_CUST(i) & vbCrLf
			END IF
		NEXT
		
		DATA_TIME = "資料時間：" & MID(AR(2), 57, 20)
		
		'IF STR = "" THEN
		'	ans = MSGBOX ("資料時間：" & MID(AR(2), 57, 20) & "  是否繼續?", vbYesNo)
		'ELSE
		IF STR <> "" THEN
			ans = MSGBOX ("資料時間：" & MID(AR(2), 57, 20) & vbCrLf & vbCrLf  & "外搬例外訂單客戶:" & vbCrLf & vbCrLf & STR & vbCrLf &vbTab &vbTab & "是否繼續?", vbYesNo)
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

              '庫別
                 RS.FIELDS(18).VALUE = LEFT(BR(16),2)

              '疊厚
	    IF LEFT(BR(2),1) ="E" OR LEFT(BR(2),1) ="F" THEN

                IF BR(0) = " " THEN

                 RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10)) + 76

                ELSE  RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10))
                
                END IF

            ELSE  RS.FIELDS(19).VALUE=  NUM(BR(5)) * NUM(BR(10))

            END IF

              '疊重
                 RS.FIELDS(20).VALUE=  NUM(BR(10)) * NUM(BR(15))

              '放行重
                 RS.FIELDS(21).VALUE=  NUM(BR(12)) * NUM(BR(15))

              '可出貨重
	    IF BR(3)<>"      " THEN	
		 RS.FIELDS(22).VALUE = NUM(BR(12)) * NUM(BR(15))
	    ELSEIF BR(3)="      " THEN 
		 RS.FIELDS(22).VALUE = 0
	    END IF
	    
   
              '訂單別	    
	    IF LEFT(BR(2),1) ="E" OR LEFT(BR(2),1) ="F" OR LEFT(BR(2),1) ="Q" THEN
			 RS.FIELDS(23).VALUE = "外銷"

	       ELSEIF  LEFT(BR(2),1) ="L" OR LEFT(BR(2),1) ="D" OR LEFT(BR(2),1) ="J" THEN
	    	 RS.FIELDS(23).VALUE = "內銷"

	  	   ELSEIF  LEFT(BR(2),2) ="TP" THEN 
	    	 RS.FIELDS(23).VALUE = "TP"

 	 	   ELSEIF  LEFT(BR(2),1) ="T" THEN 
	    	 RS.FIELDS(23).VALUE = "內部"

	       ELSE RS.FIELDS(23).VALUE = "其他"
	 
	    END IF 
	 
            '交月
	       RS.FIELDS(24).VALUE =  LEFT(BR(14),2)
             
            '超長重
	    IF NUM(BR(7)) > 13000 THEN
		RS.FIELDS(25).VALUE =  NUM(BR(10)) * NUM(BR(15))
	       ELSE
		RS.FIELDS(25).VALUE =  0
	    END IF

            '超寬重
	    IF NUM(BR(6)) > 3000 THEN
		RS.FIELDS(26).VALUE =  NUM(BR(10)) * NUM(BR(15))
	       ELSE
		RS.FIELDS(26).VALUE =  0
	    END IF


              '切邊&軋邊 
              RS.FIELDS(27).VALUE = RIGHT(BR(18),1)

              '鋼種 
	      RS.FIELDS(28).VALUE = TRIM(LEFT(BR(19)  ,13))
	    
	      '內銷訂單客戶 
              RS.FIELDS(29).VALUE = LEFT(BR(20),10)

              '外銷下游客戶 
	      RS.FIELDS(30).VALUE = LEFT(BR(21),6)
               
              'LY分類
       
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

            '吊移月
	       RS.FIELDS(32).VALUE =  LEFT(BR(9),2)

            'ORD訂單前7碼
	       RS.FIELDS(33).VALUE =  LEFT(BR(2),7)

            RS.MOVENEXT
        NEXT
    END SUB

    SUB PRCS_DATA
        SQL = "Select loc, 提單, sum(單重) as 重量, sum(片1) as 片數 INTO 儲區_各提單 from IA73 group by loc, 提單" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select 庫別, sum(疊重)/1000 as 內銷總重 INTO 儲區_內銷存量 from IA73 WHERE 訂單別 = '內銷' AND 庫別  IN( '01','07','17') AND OP = ' ' group by 庫別" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select 庫別, sum(疊重)/1000 as 外銷總重 INTO 儲區_外銷存量 from IA73 WHERE 訂單別 = '外銷' AND 庫別  IN( '01','07','17') AND OP = ' ' group by 庫別" 
        CONN.EXECUTE(SQL)
        
        SQL = "Select 庫別, sum(疊重)/1000 as 內部總重 INTO 儲區_內部存量 from IA73 WHERE 訂單別 = '內部' AND 庫別  IN( '01','07','17') AND OP = ' ' group by 庫別" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, MAX(厚) as 厚MAX, MIN(厚) as 厚MIN,  MAX(寬) as 寬MAX, MIN(寬) as 寬MIN,sum(疊重)/1000 as 總重,sum(放行重)/1000 as 放行 INTO 儲區_LY庫存 from IA73 WHERE 庫別  IN( '01','07','17') group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as LY庫存重 INTO 儲區_LY重 from IA73 WHERE 庫別  IN( '01','07','17') AND OP <> ' ' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 一般LY重 INTO 儲區_一般LY from IA73 WHERE 庫別  IN( '01','07','17') AND OP = '?' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 超長寬LY重 INTO 儲區_超長寬LY from IA73 WHERE 庫別  IN( '01','07','17') AND OP = 'W' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 中高碳LY重 INTO 儲區_中高碳LY from IA73 WHERE 庫別  IN( '01','07','17') AND OP = 'C' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 特殊LY重 INTO 儲區_特殊LY from IA73 WHERE 庫別  IN( '01','07','17') AND OP = 'H' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 高強度重 INTO 儲區_高強度LY from IA73 WHERE 庫別  IN( '01','07','17') AND OP = 'E' group by loc" 
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE 儲區_LY庫存 ADD LY庫存 DOUBLE, 一般 DOUBLE, 超長寬 DOUBLE, 中高碳 DOUBLE, 特殊 DOUBLE, 高強度 DOUBLE"
        CONN.EXECUTE(SQL)

        SQL = "Select 庫別, sum(疊重)/1000 as 儲區總重量, sum(放行重)/1000 as 放行總重量, sum(可出貨重)/1000 as 可出貨重量  INTO 儲區_各庫存量 from IA73 WHERE  庫別  IN( '01','07','17') group by 庫別" 
        CONN.EXECUTE(SQL)
       
        'SQL = "Select DISTINCT loc, 客戶, 提單 INTO 儲區_提單清單 from IA73  group by loc,  客戶, 提單 " 
        'CONN.EXECUTE(SQL)
		
		SQL = "Select loc, 客戶, 提單, 訂單客戶 INTO 儲區_提單清單 from IA73 group by loc, 客戶, 提單, 訂單客戶"
		CONN.EXECUTE(SQL)

        SQL = "Select DISTINCT loc, 客戶 INTO 儲區_客戶清單 from IA73  group by loc,  客戶 " 
        CONN.EXECUTE(SQL)
		
		'SQL = "Select DISTINCT loc, 訂單客戶 INTO 儲區_訂單客戶清單 from IA73  group by loc, 訂單客戶" 
        'CONN.EXECUTE(SQL)

        sql = " select loc, count(提單) as 提單張數 into 儲區_提單張數 FROM 儲區_各提單 GROUP BY loc "
        CONN.EXECUTE(SQL)
   
        SQL = "Select loc, sum(單重)/1000 as 已開提單_重量, sum(片1) as 已開提單_片數 INTO 儲區_提單總重  from IA73 WHERE 提單 <> '      ' group by loc  " 
        CONN.EXECUTE(SQL)

        SQL = "Select loc, sum(疊重)/1000 as 儲區總重量, sum(片1) as 片數, sum(疊厚) as 儲位厚 , MAX(厚) as 厚MAX, MIN(厚) as 厚MIN, sum(可出貨重)/1000 as 可出貨重 , MAX(長) as 長MAX, MIN(長) as 長MIN, MAX(交期) as 交期MAX, MIN(交期) as 交期MIN INTO 儲區_重量 from IA73 WHERE 庫別 IN('01', '07', '17', '04') group by loc " 
        CONN.EXECUTE(SQL)

        SQL = " SELECT 儲區_重量.*, 儲區_提單張數.提單張數 INTO 儲區一 FROM 儲區_重量 LEFT JOIN 儲區_提單張數 ON 儲區_重量.LOC = 儲區_提單張數.LOC "
        CONN.EXECUTE(SQL)

        SQL = " SELECT  儲區一.*, 儲區_提單總重.已開提單_重量, 儲區_提單總重.已開提單_片數 INTO 儲區二 FROM 儲區一 LEFT JOIN 儲區_提單總重 ON 儲區一.LOC =儲區_提單總重.LOC"
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE 儲區二 ADD 全部提單 VARCHAR(250)"
        CONN.EXECUTE(SQL)

        'SQL = "ALTER TABLE 儲區_重量 ADD 全部交運客戶 VARCHAR(250)"
		SQL = "ALTER TABLE 儲區_重量 ADD 全部交運客戶 VARCHAR(250), 全部訂單客戶 VARCHAR(250)"
        CONN.EXECUTE(SQL)
		
		'SQL = "ALTER TABLE 儲區_重量 ADD 全部訂單客戶 VARCHAR(250)"
        'CONN.EXECUTE(SQL)

    
    END SUB

    SUB PRCS_BL
        SQL = "ALTER TABLE 儲區_提單清單  ADD 備註 VARCHAR(11) "
        CONN.EXECUTE(SQL)

        SQL = "ALTER TABLE 儲區_提單清單  ADD 備註1 VARCHAR(11) "
        CONN.EXECUTE(SQL)
		
		SQL = "ALTER TABLE 儲區_提單清單  ADD 備註2 VARCHAR(11) "
        CONN.EXECUTE(SQL)
        
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_提單清單 "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
           IF LOC_A <> TS.FIELDS("LOC").VALUE THEN
                CS = MID(TS.FIELDS("客戶").VALUE, 2, 2)
				OCS = MID(TS.FIELDS("訂單客戶").VALUE, 2, 4)
                TS.FIELDS("備註").VALUE = CS & "->" & TS.FIELDS("提單").VALUE & " "
	            TS.FIELDS("備註1").VALUE = "< " &CS & " >" 
				'TS.FIELDS("備註2").VALUE = "<" & OCS & ">"
            ELSEIF  CS  <> MID(TS.FIELDS("客戶").VALUE, 2, 2)  THEN
                CS = MID(TS.FIELDS("客戶").VALUE, 2, 2)
				OCS = MID(TS.FIELDS("訂單客戶").VALUE, 2, 4)
                TS.FIELDS("備註").VALUE = CS & "->" & TS.FIELDS("提單").VALUE & " "
		        TS.FIELDS("備註1").VALUE =  "< " &CS & " >"
				'TS.FIELDS("備註2").VALUE = "<" & OCS & ">"
            ELSE
                TS.FIELDS("備註").VALUE = TS.FIELDS("提單").VALUE & ""
                TS.FIELDS("備註1").VALUE = ""
				'TS.FIELDS("備註2").VALUE = ""
            END IF
            LOC_A = TS.FIELDS("LOC").VALUE
           TS.MOVENEXT
        LOOP
		ON ERROR RESUME NEXT
		
		set Ts = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_提單清單 "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
            IF LOC_A <> TS.FIELDS("LOC").VALUE THEN
				OCS = MID(TS.FIELDS("訂單客戶").VALUE, 2, 4)
				TS.FIELDS("備註2").VALUE = "<" & OCS & ">"
            ELSEIF OCS <> MID(TS.FIELDS("訂單客戶").VALUE, 2, 4) THEN
                OCS = MID(TS.FIELDS("訂單客戶").VALUE, 2, 4)       
		        TS.FIELDS("備註2").VALUE = "<" & OCS & ">"
            ELSE
                TS.FIELDS("備註2").VALUE = ""
            END IF
			'MsgBox(TS.FIELDS("備註2").VALUE)
            LOC_A = TS.FIELDS("LOC").VALUE
            TS.MOVENEXT
        LOOP
		ON ERROR RESUME NEXT

        SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_提單清單 "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
            SQL = " UPDATE 儲區_重量 SET 全部交運客戶 = 全部交運客戶 & '" & TS.FIELDS("備註1").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
            CONN.EXECUTE(SQL)
			SQL = " UPDATE 儲區_重量 SET 全部訂單客戶 = 全部訂單客戶 & '" & TS.FIELDS("備註2").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
            CONN.EXECUTE(SQL)
            TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_LY重 "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET LY庫存 =  '" & TS.FIELDS("LY庫存重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_一般LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET 一般 =  '" & TS.FIELDS("一般LY重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_超長寬LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET 超長寬 =  '" & TS.FIELDS("超長寬LY重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT

   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_中高碳LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET 中高碳 =  '" & TS.FIELDS("中高碳LY重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_特殊LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET 特殊 =  '" & TS.FIELDS("特殊LY重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


   SET TS = NOTHING
        set Ts = CreateObject("ADODB.recordset") 
        sql = "select * from 儲區_高強度LY "
        Ts.open sql,conn,1,3
        DO WHILE NOT TS.EOF
              SQL = " UPDATE 儲區_LY庫存 SET 高強度 =  '" & TS.FIELDS("高強度重").VALUE  &  "' WHERE LOC = '" & TS.FIELDS("LOC").VALUE & "'"
           CONN.EXECUTE(SQL) 
           TS.MOVENEXT
        LOOP
	ON ERROR RESUME NEXT


    END SUB


    SUB XLS
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from 儲區_重量 "
         Ts.open sql,conn,1,3
         SET XL = CREATEOBJECT("EXCEL.APPLICATION")
         XL.VISIBLE = TRUE
		 WS.AppActivate  XL
         XL.WORKBOOKS.ADD
		XL.Sheets.Add

   		 XL.Activesheet.name=("儲位彙總")

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

		XL.CELLS(1,10).VALUE="庫"
		XL.CELLS(1,11).VALUE="Column"
		XL.CELLS(1,12).VALUE="Row"
		XL.CELLS(1,13).VALUE="型態"
		XL.CELLS(1,14).VALUE="訂單別"
		XL.CELLS(1,15).VALUE="可出貨"
	
'MSGBOX ""

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("儲位彙總").UsedRange.Rows.Count
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
			' XL.CELLS(K,13).VALUE="=IF(AND(LEN(R"& K & ")>6,H"& K & "<9001),""定尺短"",IF(AND(H" & K & "<13001,I" & K & ">9000,LEN(R" & K & ")>6),""定尺長"",""""))"
		' END IF
		
		XL.CELLS(K,13).VALUE="=IF(AND(LEN(R"& K & ")>6,H"& K & "<9001),""定尺短"",IF(AND(H" & K & "<13001,I" & K & ">9000,LEN(R" & K & ")>6),""定尺長"",""""))"
		
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
         XL.Sheets("儲位彙總").Select

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("儲位彙總").UsedRange.Rows.Count
		
		XL.CELLS(K,14).VALUE="=IF(COUNTIFS(IA73!Q:Q,儲位彙總!A" & K & ",IA73!X:X,""外銷"")>0,""外銷"","""")"
		
	NEXT

    XL.Sheets("各庫存量").Select

   	XL.CELLS(1,5).VALUE = "內銷總重"
	XL.CELLS(2,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""內銷"",IA73!A:A,"" "")/1000"
	XL.CELLS(3,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""內銷"",IA73!A:A,"" "")/1000"
	XL.CELLS(4,5).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""內銷"",IA73!A:A,"" "")/1000"

   	XL.CELLS(1,6).VALUE = "外銷總重"
	XL.CELLS(2,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""外銷"",IA73!A:A,"" "")/1000"
	XL.CELLS(3,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""外銷"",IA73!A:A,"" "")/1000"
	XL.CELLS(4,6).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""外銷"",IA73!A:A,"" "")/1000"

   	XL.CELLS(1,7).VALUE = "內銷可出"
	XL.CELLS(2,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""01"",IA73!A:A,"" "",IA73!X:X,""內銷"")/1000"
	XL.CELLS(3,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""07"",IA73!A:A,"" "",IA73!X:X,""內銷"")/1000"
	XL.CELLS(4,7).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""17"",IA73!A:A,"" "",IA73!X:X,""內銷"")/1000"

   	XL.CELLS(1,8).VALUE = "外銷可出"
	XL.CELLS(2,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""01"",IA73!A:A,"" "",IA73!X:X,""外銷"")/1000"
	XL.CELLS(3,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""07"",IA73!A:A,"" "",IA73!X:X,""外銷"")/1000"
	XL.CELLS(4,8).VALUE = "=SUMIFS(IA73!W:W,IA73!S:S,""17"",IA73!A:A,"" "",IA73!X:X,""外銷"")/1000"
	
   	XL.CELLS(1,10).VALUE = "風電總重"
	XL.CELLS(2,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""內銷"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(3,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""07"",IA73!X:X,""內銷"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(4,10).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""17"",IA73!X:X,""內銷"",IA73!AC:AC,""S355M*"",IA73!AF:AF,"""")/1000"
	XL.CELLS(5,10).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
	XL.CELLS(6,10).FormulaR1C1 = "=R[-1]C/R5C2"
	XL.CELLS(6,10).Style = "Percent"

   	XL.CELLS(1,11).VALUE = "定尺長"
	XL.CELLS(2,11).VALUE = "=SUMIFS(儲位彙總!B:B,儲位彙總!M:M,""定尺長"",儲位彙總!J:J,""01"",儲位彙總!N:N,"""")"

	XL.CELLS(1,12).VALUE = "定尺短"
	XL.CELLS(2,12).VALUE = "=SUMIFS(儲位彙總!B:B,儲位彙總!M:M,""定尺短"",儲位彙總!J:J,""01"",儲位彙總!N:N,"""")"

	XL.CELLS(3,13).VALUE = "外銷17"
	XL.CELLS(4,13).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""外銷"",IA73!A:A,"" "",IA73!H:H,""<9001"",IA73!C:C,""<>QJ*"")/1000"

	XL.CELLS(1,13).VALUE = "外銷07"
	XL.CELLS(2,13).VALUE = "=SUMIFS(IA73!U:U,IA73!S:S,""01"",IA73!X:X,""外銷"",IA73!A:A,"" "",IA73!H:H,""<13000"")/1000-M4"

	XL.CELLS(3,14).VALUE = "=""各庫庫存：01庫：""& ROUND(B2,0) & ""噸、07庫："" & ROUND(B3,0) & ""噸、17庫："" & ROUND(B4,0) & ""噸，總庫存："" & ROUND(B5,0) & ""噸。"""
	'XL.CELLS(4,14).VALUE = "=""   待外搬量：07庫(內銷"" & ROUND(K2,0) & ""噸+外銷"" & ROUND(M2,0) & ""噸)、17庫("" & ROUND(L2,0) & ""噸)。"""
	XL.CELLS(4,14).VALUE = "=""   待外搬量：07庫(內銷"" & ROUND(K2,0) & ""噸+外銷"" & ROUND(M2,0) & ""噸)、17庫(內銷"" & ROUND(L2,0) & ""噸+外銷"" & ROUND(M4,0) & ""噸)。"""
	'XL.CELLS(5,14).VALUE = "=""   風電庫存：01庫："" & ROUND(J2,0) & ""噸、07庫："" & ROUND(J3,0) & ""噸、17庫："" & ROUND(J4,0) & ""噸。"""
	
	STR = ""
	IF XL.CELLS(2,10).VALUE > 0 THEN
	    STR = "=""   風電庫存：01庫："" & ROUND(J2,0) & ""噸"
	    IF XL.CELLS(3,10).VALUE > 0 THEN
		    STR = STR & "、07庫："" & ROUND(J3,0) & ""噸"
		    IF XL.CELLS(4,10).VALUE > 0 THEN
			    STR = STR & "、17庫："" & ROUND(J4,0) & ""噸"
			END IF
		ELSEIF XL.CELLS(4,10).VALUE > 0 THEN
		    STR = STR & """、17庫："" & ROUND(J4,0) & ""噸"
		END IF
	ELSEIF XL.CELLS(3,10).VALUE > 0 THEN
	    STR = "=""   風電庫存：07庫："" & ROUND(J3,0) & ""噸"
		IF XL.CELLS(7,10).VALUE > 0 THEN
		    STR = STR & "、17庫："" & ROUND(J4,0) & ""噸"
		END IF
	ELSEIF  XL.CELLS(4,10).VALUE > 0 THEN
	    STR = "=""   風電庫存：17庫："" & ROUND(J4,0) & ""噸"
	END IF
	
	IF XL.CELLS(2,10).VALUE > 0 OR XL.CELLS(3,10).VALUE > 0 OR XL.CELLS(4,10).VALUE > 0 THEN
		XL.CELLS(5,14).VALUE = STR & "。"""
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
         
    XL.Sheets("各庫存量").Select

'各月交期庫存
        'XL.Cells(10, 1).Select
        
        XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "各庫存量!R8C1", "樞紐分析表10", 3
      
    With XL.ActiveSheet.PivotTables("樞紐分析表10").PivotFields("交月")
        .Orientation = 1
        .Position = 1
    End With
    With XL.ActiveSheet.PivotTables("樞紐分析表10").PivotFields("庫別")
        .Orientation = 2
        .Position = 1
    End With

   XL.ActiveSheet.PivotTables("樞紐分析表10").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表10").PivotFields("疊重"), "加總 - 疊重",  -4157

   XL.ActiveSheet.PivotTables("樞紐分析表10").PivotFields("加總 - 疊重").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("樞紐分析表10").CompactLayoutColumnHeader = "庫別"
    XL.ActiveSheet.PivotTables("樞紐分析表10").CompactLayoutRowHeader = "交月"
    XL.ActiveSheet.PivotTables("樞紐分析表10").DataPivotField.PivotItems("加總 - 疊重").Caption = "庫存重"

    With XL.ActiveSheet.PivotTables("樞紐分析表10").PivotFields("庫別")
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

 XL.ActiveSheet.PivotTables("樞紐分析表10").PivotSelect "", 0, True
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

'各吊移月庫存

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "各庫存量!R23C1", "樞紐分析表13", 3
   

    With XL.ActiveSheet.PivotTables("樞紐分析表13").PivotFields("庫別")
        .Orientation = 2
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表13").PivotFields("吊移月")
        .Orientation = 1
        .Position = 1
    End With

    XL.ActiveSheet.PivotTables("樞紐分析表13").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表13").PivotFields("疊重"), "加總 - 疊重",  -4157

    XL.ActiveSheet.PivotTables("樞紐分析表13").PivotFields("加總 - 疊重").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("樞紐分析表13").CompactLayoutColumnHeader = "庫別"


    XL.ActiveSheet.PivotTables("樞紐分析表13").CompactLayoutRowHeader = "吊移月"
    XL.ActiveSheet.PivotTables("樞紐分析表13").DataPivotField.PivotItems("加總 - 疊重").Caption = "吊移重"


   With XL.ActiveSheet.PivotTables("樞紐分析表13").PivotFields("庫別")
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

 XL.ActiveSheet.PivotTables("樞紐分析表13").PivotSelect "", 0, True
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

'內部訂單各月交期庫存
   XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C33", 3).CreatePivotTable "各庫存量!R29C8", "樞紐分析表11", 3
      
    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("交月")
        .Orientation = 1
        .Position = 1
    End With
    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("庫別")
        .Orientation = 2
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("訂單別")
        .Orientation = 3
        .Position = 1
    End With

  With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("OP")
        .Orientation = 3
        .Position = 2
    End With
 
    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("訂單別")
        .PivotItems("TP").Visible = False
        .PivotItems("內銷").Visible = False
        .PivotItems("外銷").Visible = False
        .PivotItems("其他").Visible = False
        .PivotItems("(blank)").Visible = False

    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("OP")
ON ERROR RESUME NEXT
        .PivotItems("?").Visible = False
        .PivotItems("C").Visible = False
        .PivotItems("H").Visible = False
        .PivotItems("R").Visible = False
        .PivotItems("S").Visible = False
        .PivotItems("W").Visible = False
        .PivotItems("E").Visible = False
    End With

   XL.ActiveSheet.PivotTables("樞紐分析表11").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("疊重"), "加總 - 疊重",  -4157

   XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("加總 - 疊重").NumberFormat = "#,##0,"

   XL.ActiveSheet.PivotTables("樞紐分析表11").CompactLayoutColumnHeader = "庫別"
    XL.ActiveSheet.PivotTables("樞紐分析表11").CompactLayoutRowHeader = "交月"
    XL.ActiveSheet.PivotTables("樞紐分析表11").DataPivotField.PivotItems("加總 - 疊重").Caption = "內領重"
   
    With XL.ActiveSheet.PivotTables("樞紐分析表11").PivotFields("庫別")
        .PivotItems("02").Visible = False '550
        .PivotItems("04").Visible = False
        .PivotItems("H0").Visible = False
        .PivotItems("S0").Visible = False
        .PivotItems("T1").Visible = False
        .PivotItems("Q0").Visible = False
        .PivotItems("M1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

    XL.ActiveSheet.PivotTables("樞紐分析表11").PivotSelect "", 0, True
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


    'LEEWAY分類表

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C32", 3).CreatePivotTable "各庫存量!R8C8", "樞紐分析表12", 3
    'XL.Sheets("各庫存量").Select


    With XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("LY分類")
        .Orientation = 1
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("庫別")
        .Orientation = 2
        .Position = 1
    End With

    XL.ActiveSheet.PivotTables("樞紐分析表12").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("疊重"), "加總 - 疊重", -4157

    XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("加總 - 疊重").NumberFormat = "#,##0,"

    With XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("庫別")

    XL.ActiveSheet.PivotTables("樞紐分析表12").DataPivotField.PivotItems("加總 - 疊重").Caption = "LY重"

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

    With XL.ActiveSheet.PivotTables("樞紐分析表12").PivotFields("LY分類")
        .PivotItems("(blank)").Visible = False
    End With


    XL.ActiveSheet.PivotTables("樞紐分析表12").PivotSelect "", 0, True
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

   		XL.Activesheet.name=("儲位使用")

		XL.ActiveWorkbook.PivotCaches.Create (1,"儲位彙總!R1C1:R1000C15", 3).CreatePivotTable "儲位使用!R3C1", "樞紐分析表4", 3
  
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").Orientation = 3
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").Position = 1

		XL.ActiveSheet.PivotTables("樞紐分析表4").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("儲區總重量"), "計數 - 片數",  -4112

		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("Row").Orientation =1
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("Row").Position = 1

		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("Column").Orientation = 2
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("Column").Position = 1
 
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").PivotItems("07").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").PivotItems("17").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").PivotItems("(blank)").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").PivotItems("04").Visible = False
 
		XL.ActiveSheet.PivotTables("樞紐分析表4").PivotFields("庫").EnableMultiplePageItems = True

		' XL.ActiveWorkbook.PivotCaches.Create (1,"儲位彙總!R1C1:R1000C15", 3).CreatePivotTable "儲位使用!R16C1", "樞紐分析表5", 3
  
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").Orientation = 3
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").Position = 1

		' XL.ActiveSheet.PivotTables("樞紐分析表5").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("儲區總重量"), "計數 - 片數",  -4112

		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("Row").Orientation =1
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("Row").Position = 1

		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("Column").Orientation = 2
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("Column").Position = 1
 
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").PivotItems("01").Visible = False
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").PivotItems("17").Visible = False
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").PivotItems("(blank)").Visible = False
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").PivotItems("04").Visible = False
 
		' XL.ActiveSheet.PivotTables("樞紐分析表5").PivotFields("庫").EnableMultiplePageItems = True

		XL.ActiveWorkbook.PivotCaches.Create (1,"儲位彙總!R1C1:R1000C15", 3).CreatePivotTable "儲位使用!R32C1", "樞紐分析表6", 3
  
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").Orientation = 3
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").Position = 1

		XL.ActiveSheet.PivotTables("樞紐分析表6").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("儲區總重量"), "計數 - 片數",  -4112

		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("Row").Orientation =1
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("Row").Position = 1

		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("Column").Orientation = 2
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("Column").Position = 1
 
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").PivotItems("01").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").PivotItems("07").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").PivotItems("(blank)").Visible = False
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").PivotItems("04").Visible = False
 
		XL.ActiveSheet.PivotTables("樞紐分析表6").PivotFields("庫").EnableMultiplePageItems = True
		XL.Cells.Select
		XL.Selection.ColumnWidth = 2.50


     END SUB
    
    SUB WHI
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from 儲區_各庫存量 "
         Ts.open sql,conn,1,3

		XL.Sheets.Add

   		XL.Activesheet.name=("各庫存量")
   		
   		FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts
                  
         SET Ts = NOTHING
         
         
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select 內銷總重 from 儲區_內銷存量 "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+7).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         
         XL.CELLS(2,7).COPYFROMRECORDSET Ts
         SET Ts = NOTHING
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select 外銷總重 from 儲區_外銷存量 "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+8).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         XL.CELLS(2,8).COPYFROMRECORDSET Ts
         SET Ts = NOTHING
         
         set Ts = CreateObject("ADODB.recordset") 
         sql = "select 內部總重 from 儲區_內部存量 "
         Ts.open sql,conn,1,3
         
         FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+9).VALUE = Ts.FIELDS(I).NAME
         NEXT 
         
         XL.CELLS(2,9).COPYFROMRECORDSET Ts
  



        XL.Cells.Select	
        XL.Selection.NumberFormatLocal = "#,##0_ "
        XL.Cells.EntireColumn.AutoFit
        
           FOR J = 2 TO XL.ActiveWorkbook.Worksheets("各庫存量").UsedRange.Columns.Count       
           XL.CELLS(5,J).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
		   XL.CELLS(6,J).FormulaR1C1 = "=R[-1]C/R5C2"
		   XL.CELLS(6,J).Style = "Percent"
     
         NEXT
         
   	END SUB


SUB CHART
   	WITH XL
   		.Sheets("儲位彙總").Select
		.Range("A:A,D:D").Select
		.Charts.Add  '儲位厚
                .Sheets("Chart1").Name = "儲位厚"
   	END WITH
   	
END SUB


SUB LYLOC

        set Ts = CreateObject("ADODB.recordset") 
         sql = "select * from 儲區_LY庫存 "
         Ts.open sql,conn,1,3

		XL.Sheets.Add

   		XL.Activesheet.name=("LY儲位")
   		
   	 FOR I = 0 TO Ts.FIELDS.COUNT -1
             XL.CELLS(1, I+1).VALUE = Ts.FIELDS(I).NAME
         NEXT 

         XL.CELLS(2,1).COPYFROMRECORDSET Ts
                  
         SET Ts = NOTHING

		XL.CELLS(1,14).VALUE  ="放行比"
		XL.CELLS(1,15).VALUE = "LY比"

	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("LY儲位").UsedRange.Rows.Count
		XL.CELLS(K,14).VALUE="=G" & K & "/F" & K
		XL.CELLS(K,15).VALUE="=H" & K & "/F" & K
	NEXT

		XL.CELLS(1,16).VALUE  ="厚度"
		XL.CELLS(1,17).VALUE = "寬度"
		XL.CELLS(1,18).VALUE = "分類"
		XL.CELLS(1,19).VALUE = "AP板"
		XL.CELLS(1,20).VALUE = "PX板"


	FOR K = 2 TO XL.ActiveWorkbook.Worksheets("LY儲位").UsedRange.Rows.Count

                XL.CELLS(K,16).VALUE="=IF(B" & K & "<=12.7,"""",IF(AND(B" & K & "<=28,C" & K & ">12.7),"""",IF(C" & K & ">28,"""",IF(LEFT(A" & K & ",2)=""01"","""",""厚異常""))))"

                XL.CELLS(K,17).VALUE="=IF(LEFT(A" & K & ",2)<>""01"",IF(D" & K & ">3250,""寬異常"",""""),"""")"

                XL.CELLS(K,18).VALUE="=IF(AND(COUNTBLANK(I" & K & ":M" & K & ")<4,LEFT(A" & K & ",2)<>""01""),""混儲"","""")"

                XL.CELLS(K,19).VALUE="=IF(COUNTIFS(IA73!Q:Q,LY儲位!A" & K & ",IA73!AC:AC,""APPLY HEAT"")>0,""AP板 *""& COUNTIFS(IA73!Q:Q,LY儲位!A" & K & ",IA73!AC:AC,""APPLY HEAT""),"""")"            

                XL.CELLS(K,20).VALUE="=IF(COUNTIFS(IA73!Q:Q,LY儲位!A" & K & ",IA73!AC:AC,""PX1"")>0,""PX板*""& COUNTIFS(IA73!Q:Q,LY儲位!A" & K & ",IA73!AC:AC,""PX1""),"""")"            
	NEXT
	
  	NN=XL.ActiveWorkbook.Worksheets("LY儲位").UsedRange.Rows.Count

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

   		XL.Activesheet.name=("內部訂單")

    '各單位內部訂單清單

    XL.ActiveWorkbook.PivotCaches.Create(1, "IA73!R1C1:R1048576C34", 3).CreatePivotTable "內部訂單!R6C1", "樞紐分析表14", 3
   
    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("ORD")
        .Orientation = 1
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("客戶")
        .Orientation = 2
        .Position = 1
    End With


    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("OP")
        .Orientation = 3
        .Position = 1
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("訂單別")
        .Orientation = 3
        .Position = 2
    End With

    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("庫別")
        .Orientation = 3
        .Position = 3
    End With




 
    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("訂單別")
	ON ERROR RESUME NEXT
        .PivotItems("TP").Visible = False
        .PivotItems("內銷").Visible = False
        .PivotItems("外銷").Visible = False
        .PivotItems("其他").Visible = False
        .PivotItems("(blank)").Visible = False

    End With

    XL.ActiveSheet.PivotTables("樞紐分析表14").AddDataField XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("疊重"), "加總 - 疊重", -4157

    XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("加總 - 疊重").NumberFormat = "#,##0,"

    XL.ActiveSheet.PivotTables("樞紐分析表14").DataPivotField.PivotItems("加總 - 疊重").Caption = "內部重"

    With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("庫別")

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

 With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("OP")
ON ERROR RESUME NEXT
        .PivotItems("?").Visible = False
        .PivotItems("C").Visible = False
        .PivotItems("H").Visible = False
        .PivotItems("R").Visible = False
        .PivotItems("S").Visible = False
        .PivotItems("W").Visible = False
        .PivotItems("E").Visible = False
    End With


  '  With XL.ActiveSheet.PivotTables("樞紐分析表14").PivotFields("LY分類")
  '      .PivotItems("(blank)").Visible = False
  '  End With


    XL.ActiveSheet.PivotTables("樞紐分析表14").PivotSelect "", 0, True
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

 XL.Sheets("各庫存量").Select
END SUB



END CLASS