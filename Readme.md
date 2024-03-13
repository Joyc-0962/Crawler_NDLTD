## 台灣博碩士論文網爬蟲 NDLTD
---
這篇簡單記錄如何從一個已經列出來的指導教授列表，來爬取這個指導教授指導過的學生畢業論文相關資訊。

博碩士論文網會記錄你的 Session 資訊，因此當我們把連結的網址給別人時，別人並沒有辦法看到我們轉貼的文章，只會被重新導回首頁，同樣的邏輯，我們如果只是單純的 get 特定的網址也沒辦法取得需要的資訊，所以我們要先開一個Session，並 post 查詢的參數到對方的伺服器，讓對方記得我們，然後再用這個 Session 去 get 我們需要的資料。

原本應該是這樣的...
但是不知道為什麼會有cookie的問題，或是憑證有時限的關係，都會被重新導向首頁。


後來的解決方案就是完全透過 Selenium 來爬博碩士論文網的文章，超級痛苦ＱＡＱ
因為台灣博碩士論文網爬蟲會鎖掉IP，記得打開 proxy server，不然自己的IP就進黑名單拉～
```
python open_web_title.py
```
output: 

    root
    │── all_professor_title.csv
    │── done_professor_title.csv   
    │── blank_professor_title.csv 
    └── log.txt
執行後會產生四隻檔案
- all_professor_title.csv -> 所有爬蟲資料
- done_professor_title.csv -> 所有爬過的教授名單
- blank_professor_title.csv -> 沒有資料的教授名單
- log_title.txt

## 學術學術研發服務網爬蟲 NSTC
---
這篇簡單記錄如何從一個已經列出來的指導教授列表，來爬取這個教授的計畫及著作。
```
python NSTC.py
```
output: 
    root
    │── nstc_professor_info.csv
    │── nstc_projects_list.csv   
    │── nstc_work_list.csv   
    │── nstc_done_professor_title.csv   
    │── nstc_blank_professor_title.csv 
    └── log_nstc.txt
執行後會產生六隻檔案
- nstc_professor_info.csv -> 教授的個人資料
- nstc_projects_list.csv -> 教授的計畫
- nstc_work_list.csv -> 教授的著作
- nstc_done_professor_title.csv -> 所有爬過的教授名單
- nstc_blank_professor_title.csv -> 沒有資料的教授名單
- log.txt
