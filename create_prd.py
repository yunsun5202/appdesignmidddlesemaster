from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION

out='個人行程與任務管理系統_PRD.docx'
doc=Document()
sec=doc.sections[0]; sec.top_margin=Inches(.68);sec.bottom_margin=Inches(.65);sec.left_margin=Inches(.78);sec.right_margin=Inches(.78)
styles=doc.styles
for nm in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
    st=styles[nm];st.font.name='Microsoft JhengHei';st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');st.font.color.rgb=RGBColor(0,0,0)
styles['Normal'].font.size=Pt(9.5);styles['Normal'].paragraph_format.space_after=Pt(5)
styles['Title'].font.size=Pt(25);styles['Title'].font.bold=True;styles['Title'].paragraph_format.space_after=Pt(5)
styles['Heading 1'].font.size=Pt(15);styles['Heading 1'].font.bold=True;styles['Heading 1'].paragraph_format.space_before=Pt(14);styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].font.size=Pt(11);styles['Heading 2'].font.bold=True;styles['Heading 2'].paragraph_format.space_before=Pt(8);styles['Heading 2'].paragraph_format.space_after=Pt(3)

def shade(cell,color):
    tcPr=cell._tc.get_or_add_tcPr();shd=OxmlElement('w:shd');shd.set(qn('w:fill'),color);tcPr.append(shd)
def cell_margins(cell,top=90,start=100,bottom=90,end=100):
    tc=cell._tc;tcPr=tc.get_or_add_tcPr();m=OxmlElement('w:tcMar')
    for name,val in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        el=OxmlElement('w:'+name);el.set(qn('w:w'),str(val));el.set(qn('w:type'),'dxa');m.append(el)
    tcPr.append(m)
def table(headers,rows,widths=None):
    t=doc.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.style='Table Grid'
    t.autofit=False
    if widths:
        for col,w in zip(t.columns,widths): col.width=Inches(w)
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i];c.text=h;shade(c,'285C43');c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(c,120,110,120,110)
        for r in c.paragraphs[0].runs:r.font.bold=True;r.font.color.rgb=RGBColor(255,255,255);r.font.size=Pt(9)
    trPr=t.rows[0]._tr.get_or_add_trPr();repeat=OxmlElement('w:tblHeader');repeat.set(qn('w:val'),'true');trPr.append(repeat)
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for i,val in enumerate(row):
            cells[i].text=str(val);cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(cells[i])
            if ri%2:shade(cells[i],'F2F6F2')
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after=Pt(0)
                for r in p.runs:r.font.size=Pt(8.5)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t
def para(text='',boldlead=None):
    p=doc.add_paragraph()
    if boldlead and text.startswith(boldlead):p.add_run(boldlead).bold=True;p.add_run(text[len(boldlead):])
    else:p.add_run(text)
    return p
def bullets(items):
    for x in items:
        p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(x)

p=doc.add_paragraph(style='Title');p.add_run('個人行程與任務管理系統')
p=doc.add_paragraph();r=p.add_run('產品需求文件 PRD');r.bold=True;r.font.size=Pt(12);r.font.color.rgb=RGBColor(40,92,67)
para('本文件定義一套支援手機與電腦的個人行程與任務管理系統，讓使用者能在每日總覽中掌握時間安排、追蹤待辦進度，並以通知降低遺漏風險。內容依目前已確認的需求與互動原型整理，供產品、設計與工程團隊作為 MVP 規劃基礎。')
table(['文件資訊','內容'],[
['產品暫名','日序 Daymark'],['文件版本','1.0'],['文件日期','2026 年 10 月 30 日'],['產品範圍','個人使用；響應式 Web／前端原型'],['目標平台','桌面瀏覽器與手機瀏覽器']],[1.55,5.75])

doc.add_heading('一 產品概述',level=1)
para('日序是一套以每日規劃為核心的個人管理工具。產品將有明確開始時間的行程與有截止日期的任務放在同一個日視圖中，同時保留未排時間任務清單。使用者可依優先級、分類與標籤管理事項，並收到重要時間點提醒。')
doc.add_heading('產品願景',level=2);para('讓使用者每天快速看清楚「今天有哪些安排、哪些事情還沒完成、下一步該做什麼」。')
doc.add_heading('問題陳述',level=2);bullets(['行程與待辦散落在不同工具，使用者難以一次掌握每日全貌。','任務有截止日但未安排進時間表，容易被忽略。','缺少可辨識的重要程度與分類方式，整理和篩選成本高。','若沒有事前提醒，使用者可能錯過行程或任務期限。'])

doc.add_heading('二 目標與成功指標',level=1)
table(['目標','衡量方式（建議）','MVP 目標'],[
['快速掌握當日安排','每日總覽可呈現當日行程、任務與未排時任務','核心內容在首頁一頁可見'],
['降低任務遺漏','使用者在期限前完成任務的比例；逾期任務比例','上線後建立基準，再依使用情況訂定改善幅度'],
['提升規劃完成率','每週至少新增／完成一筆項目的活躍使用者比例','建立事件追蹤後觀察，不預設未驗證數字'],
['提醒可用性','提醒權限開啟率、通知觸發成功率','支援通知權限、提醒排程與失敗回饋']],[1.5,3.55,2.25])
para('指標基準尚待正式產品上線後蒐集；本文件不假設尚未驗證的使用數據。')

doc.add_heading('三 目標使用者與使用情境',level=1)
table(['使用者／情境','需求'],[
['個人使用者','以手機或電腦查看今天的行程與待辦，管理個人生活、工作、學習及健康事項。'],
['開始一天前','確認當日行程時間、優先任務與待排程事項。'],
['收到提醒時','在行程開始前一小時或任務截止前一天收到提示，快速確認或更新事項。'],
['整理待辦時','透過分類與標籤找到同一主題的任務，調整優先級或完成狀態。']],[1.65,5.65])

doc.add_heading('四 產品範圍',level=1)
doc.add_heading('MVP 範圍內',level=2);bullets(['響應式每日總覽，適用桌面與手機瀏覽器。','每日時間軸：顯示當天已排時間的行程及任務，依時間排序。','未安排時間任務清單：突出顯示尚未排入時間軸且未完成的任務。','建立、編輯、刪除任務與行程；標記任務完成或恢復未完成。','任務與行程欄位：名稱、日期、時間、優先級、分類、標籤、備註及提醒開關。','優先級：高、中、低；分類：工作、個人、學習、健康、其他；允許多個自由標籤。','側邊分類展開／收合，依分類瀏覽未完成待辦；顯示分類待辦數量。','日期切換與月曆選日；資料在同一瀏覽器本機保存。','瀏覽器通知權限請求及提醒設定介面。'])
doc.add_heading('MVP 範圍外／後續評估',level=2);bullets(['帳號、登入、雲端同步、多裝置即時同步。','多人共享、協作、任務指派與權限管理。','完整週／月行事曆檢視與拖放排程。','重複行程／任務、循環提醒、附件與外部日曆整合。','跨瀏覽器／裝置推播、離線背景通知與通知伺服器排程。'])

doc.add_heading('五 核心流程',level=1)
table(['流程','步驟'],[
['新增任務','選擇新增任務 → 輸入名稱、日期／截止時間 → 設定優先級、分類、標籤與提醒 → 儲存 → 顯示於每日時間軸或未安排清單。'],
['新增行程','選擇新增行程 → 輸入名稱、日期、開始／結束時間 → 設定分類、標籤與提醒 → 儲存 → 依時間顯示於時間軸。'],
['完成任務','在時間軸或清單點選完成控制 → 更新完成狀態與每日進度 → 從未完成分類清單移除。'],
['依分類查看待辦','展開左側分類 → 查看該分類未完成任務及到期日／優先級 → 點選任務開啟編輯。'],
['切換日期','使用前／後日期控制或月曆選日 → 更新時間軸、待辦、統計與選取日期。']],[1.45,5.85])

doc.add_heading('六 功能需求',level=1)
table(['ID','需求','優先級','驗收條件'],[
['F-01','每日總覽與日期切換','Must','選取日期改變時，日期標題、時間軸、任務清單及統計同步更新。'],
['F-02','行程時間軸','Must','當日有時間的項目依開始時間排序；行程呈現開始與結束時間，任務呈現截止時間。'],
['F-03','未排時間任務','Must','顯示未完成、未設定時間的任務，呈現截止日期、分類與優先級。'],
['F-04','任務 CRUD 與狀態','Must','可新增、編輯、刪除任務；可標記完成／恢復；重新整理後資料仍保留於同一瀏覽器。'],
['F-05','行程 CRUD','Must','可新增、編輯、刪除行程；行程必須設定開始時間，可選填結束時間。'],
['F-06','優先級／分類／標籤','Must','任務與行程可設定優先級及分類；標籤支援使用者輸入多個值。'],
['F-07','分類展開清單','Must','點選分類可展開／收合該分類未完成任務，顯示筆數；點選任務可進入編輯。'],
['F-08','提醒設定','Must','任務預設截止前一天提醒；行程預設開始前一小時提醒；使用者可關閉提醒。'],
['F-09','月曆選日','Should','可切換月分並選取日期；選取日同步更新主要內容。'],
['F-10','瀏覽器通知權限','Should','使用者可主動授權通知；拒絕或不支援時提供清楚提示，不阻斷基本管理功能。']],[.62,1.55,.72,4.41])

doc.add_heading('七 資料欄位與規則',level=1)
table(['欄位','適用項目','規則'],[
['識別碼','任務／行程','系統產生且唯一。'],['名稱','任務／行程','必填；最多 80 字。'],['類型','任務／行程','任務或行程。'],['日期','任務／行程','必填；任務代表截止日期，行程代表發生日期。'],['時間','任務／行程','任務截止時間選填；行程開始時間必填。'],['結束時間','行程','選填；若提供，須晚於開始時間。'],['優先級','任務／行程','高／中／低；預設中。'],['分類','任務／行程','工作、個人、學習、健康、其他；預設工作。'],['標籤','任務／行程','零至多個自由文字標籤，使用逗號分隔輸入。'],['提醒','任務／行程','任務預設截止前一天；行程預設開始前一小時；可停用。'],['完成狀態','任務','未完成／完成；完成項目不列入側邊分類未完成清單。'],['備註','任務／行程','選填文字。']],[1.15,1.2,4.95])

doc.add_heading('八 通知與提醒規格',level=1)
bullets(['任務提醒：於截止日前一天的相同時間發送；若無截止時間，建議使用產品預設時間（待產品決策，建議上午 9:00）。','行程提醒：於開始時間前一小時發送。','僅對提醒已開啟且未刪除的項目排程通知。','若項目時間／日期變更，更新排程；項目刪除或提醒關閉時取消排程。','通知權限被拒、瀏覽器不支援或排程無法建立時，提供可理解的狀態提示。','正式產品需依平台選擇可靠的通知架構（例如服務端推播或平台原生通知）；純前端頁面無法保證在瀏覽器關閉時送達。'])

doc.add_heading('九 非功能需求',level=1)
table(['面向','需求'],[
['響應式與可用性','支援常見手機與桌面瀏覽器尺寸；主要任務操作在手機觸控下可完成。'],
['效能','主要視圖在一般網路及一般裝置上快速呈現；後端上線後需訂定量化效能目標。'],
['可存取性','表單欄位有標籤；可鍵盤操作分類展開、對話框與互動控制；顏色不是唯一狀態提示。'],
['資料持久性','原型以瀏覽器本機儲存；正式版需明確定義雲端保存、備份與資料匯出策略。'],
['隱私與安全','個人行程資料屬私人資料；正式版需採用適當驗證、授權、傳輸加密及資料保留政策。'],
['相容性','支援現行主流桌面與手機瀏覽器；通知能力依瀏覽器與作業系統支援狀況而異。']],[1.45,5.85])

doc.add_heading('十 介面與資訊架構',level=1)
table(['區域','內容與行為'],[
['左側導覽','每日總覽、週行事曆（後續）、所有任務；下方分類清單可展開查看同類未完成任務。'],
['每日摘要','完成進度、今日行程數、待完成任務數。'],
['主要時間軸','行程與有時間任務混合依時間排序；項目卡片可開啟編輯，任務可直接完成。'],
['未安排時間','未排入時間軸的待辦卡片；顯示到期日、分類及優先級。'],
['月曆','切換月份、選日期；所選日期同步至每日總覽。'],
['新增／編輯對話框','切換任務或行程，輸入日期時間、優先級、分類、標籤、提醒與備註。']],[1.45,5.85])

doc.add_heading('十一 例外情境與待決策事項',level=1)
table(['議題','目前建議／需確認'],[
['截止日只有日期、沒有時間','建議以當地時區上午 9:00 作為提醒基準，後續由產品負責人確認。'],
['任務設定時間的呈現','原型將有時間的任務排入時間軸；無時間任務顯示在未安排清單。需確認任務時間的語意是「截止時間」而非「預計開始時間」。'],
['過期任務','應持續顯示並標記逾期，避免自動隱藏；完成或刪除後才移出待辦清單。'],
['分類與自訂分類','MVP 使用固定分類及其他；是否允許使用者建立、改名、排序分類，列入後續決策。'],
['資料同步','多裝置使用是目標體驗；目前原型本機儲存並不會跨手機與電腦同步，正式版需帳號及雲端服務。'],
['通知可靠性','正式版需決定是否建置服務端提醒排程，以及瀏覽器關閉、裝置離線與時區變更時的處理規則。'],
['週檢視','導覽已有週行事曆入口，但 MVP 目前聚焦每日視圖；是否納入第一版需排定優先級。']],[1.65,5.65])

doc.add_heading('十二 版本建議',level=1)
table(['階段','建議交付'],[
['原型／需求驗證','驗證每日時間軸、未排任務、分類展開及新增編輯流程；收集使用者回饋。'],
['MVP 1.0','帳號與雲端資料、任務／行程 CRUD、每日總覽、分類／標籤、提醒排程與手機桌面響應式體驗。'],
['後續版本','週／月檢視、重複事項、自訂分類、搜尋篩選、日曆整合、共享及更完整通知能力。']],[1.5,5.8])

doc.add_heading('附錄 原型與本文件對應',level=1)
para('目前互動原型提供每日時間軸、未安排任務、新增／編輯任務與行程、完成狀態、分類與標籤欄位、月曆選日、分類展開待辦及瀏覽器通知權限示範。原型使用本機瀏覽器儲存，通知排程與雲端同步尚未實作；此處所列為後續正式產品需求，不代表目前原型已具備背景通知或多裝置同步。')

# Footer page number
for section in doc.sections:
    fp=section.footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    run=fp.add_run('日序 Daymark  |  PRD  ·  ');run.font.size=Pt(8);run.font.color.rgb=RGBColor(110,120,110)
    fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');fp._p.append(fld)

doc.save(out)
print(out)
