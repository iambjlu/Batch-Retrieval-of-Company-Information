<h1>批量查詢公司資料</h1>
可以查到公司的設立狀態、公司全名、資本額、實收資本額、負責人、公司地址、登記機關名稱、核准設立日期、最後核准變更日期...等。<br>
資料來源:商工行政資料開放平台
<hr>
<h3>main_company_name.py</h3>
<pre>使用說明:
把公司名字關鍵字放進companies.txt(與此程式檔案置於同樣目錄)
可以獲得output.csv，可以使用Excel開啟</pre>

<h3>main_company_no.py</h3>
<pre>使用說明:
把公司統編放進companies.txt(與此程式檔案置於同樣目錄)
可以獲得output.csv，可以使用Excel開啟</pre>

<hr>
<h2>使用範例</h2>
輸入 companies.txt
<pre>
22099131
53433060
20838245
03251108
24004006</pre><br>
會得到輸出 output.csv<br>
(使用Excel開啟並另存xlsx格式，可以獲得更好的使用體驗)
<pre>
Business_Accounting_NO,Company_Name,Company_Status_Desc,Capital_Stock_Amount,Paid_In_Capital_Amount,Responsible_Name,Register_Organization_Desc,Company_Location,Company_Setup_Date,Change_Of_Approval_Data,Revoke_App_Date,Case_Status,Case_Status_Desc,Sus_App_Date,Sus_Beg_Date,Sus_End_Date
22099131,台灣積體電路製造股份有限公司,核准設立,280500000000,259327332420,魏哲家,國家科學及技術委員會新竹科學園區管理局,新竹科學園區新竹市力行六路8號,0760221,1130911,,,,,,
53433060,睿能創意股份有限公司,核准設立,5200000000,4183400000,曾達夢,商業發展署,桃園市龜山區頂湖路33號,1000829,1131009,,,,,,
20838245,汎德永業汽車股份有限公司,核准設立,1000000000,807087350,唐慕蓮,經濟部商業司,臺北市內湖區行愛路100號6樓,0681107,1120719,,,,,,
03251108,和泰汽車股份有限公司,核准設立,6000000000,5571027680,黃南光,商業發展署(實收資本額5億元以上),臺北市中山區松江路121號8~14樓,0440425,1130627,,,,,,
24004006,三陽工業股份有限公司,核准設立,9500000000,7974896040,吳清源,經濟部商業司,新竹縣湖口鄉鳳山村1鄰中華路3號,0500829,1120718,,,,,,</pre>
