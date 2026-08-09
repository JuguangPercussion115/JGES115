import os
import configparser
import urllib.request
import pandas as pd
import numpy as np
import shutil
import hashlib
from datetime import datetime

# ==============================================================================
# 1. 基礎工具：網路下載與 MD5 檢查
# ==============================================================================
def DownLoadFile(url, file_name=""):   
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'} 
    try:
        request = urllib.request.Request(url, headers=headers)
        response = urllib.request.urlopen(request)
        if file_name == "":
            file_name = response.info().get_filename() or "no_filename"
        
        block_sz = 8192
        with open(file_name, 'wb') as f:
            while True:
                buffer = response.read(block_sz)
                if not buffer:
                    break
                f.write(buffer)
        print(f" [下載成功] -> {file_name}")
    except Exception as e:
        print(f" [下載錯誤] 無法下載 {file_name}: {e}")

def check_md5(filename):
    if not os.path.exists(filename):
        return ""
    try:
        pd.read_excel(filename).to_csv(filename + ".md5", index=False)
        m = hashlib.md5()
        with open(filename + ".md5", 'rb') as f:
            line = f.read()
            m.update(line)
        return m.hexdigest()
    except:
        return ""

def LoadPatten(filename, patten_tag):
    ret = ""
    if not os.path.exists(filename):
        return f"[{patten_tag} Template Not Found]"
    with open(filename, mode='r', encoding='utf-8') as f:
        content = f.read()
        f1 = content.find("#" + patten_tag.upper() + "#")
        if f1 >= 0:
            f2 = content.find("#" + patten_tag + "#", f1 + 1)
            if f2 >= 0:                
                ret = content[f1 + len(patten_tag) + 3:f2]
    return ret

# ==============================================================================
# 2. 原版 HTML 渲染引擎 (GenReport) - 100% 還原舊版網頁格式
# ==============================================================================
def GenReport(df, song_list, pgm_info):
    break_line = '''
<tr bgcolor="#000000">
<tr bgcolor="#000000">
  <td bgcolor="#000000" rowspan=1 colspan="{}">
  </td>
</tr>
</tr>
'''    
    # 讀取 INI 設定中的產出路徑與標題
    html_file_name = pgm_info.get('產出html檔名', 'html/index.html')
    temp_html_name = "ini/Sample_Race.html"
    
    html_header = LoadPatten(temp_html_name, "HEADER")
    html_header = html_header.replace("#TITLE#", pgm_info.get('產出html標題', '成績看板'))
    html_header = html_header.replace("#UPDATE_DATE#", pgm_info.get("_update_string", ""))    
    html_header = html_header.replace("#UPDATE_DATE1#", str(pgm_info.get("adj_update_time", "")))
    
    html_header1 = LoadPatten(temp_html_name, "HEADER_R1")
    html_header1_1 = LoadPatten(temp_html_name, "HEADER_R1_1")
    html_header1_2 = LoadPatten(temp_html_name, "HEADER_R1_2")
        
    html_student = LoadPatten(temp_html_name, "STUDENT_R")
    html_student_rs = LoadPatten(temp_html_name, "STUDENT_RS")
    
    max_score = df['總積分'].max() if len(df) > 0 else 100
    
    # 建立目錄確保寫入不會失敗
    os.makedirs(os.path.dirname(html_file_name), exist_ok=True)
    
    with open(html_file_name, mode='w', encoding='utf-8') as f:
        f.write(html_header)
        f.write(html_header1)
        
        for song in song_list:
            tdstr = html_header1_1.replace("#SONG_NAME#", song)
            f.write(tdstr)
        f.write("</tr>")   
        
        f.write("<tr>") 
        for song in song_list:
            f.write(html_header1_2)
        f.write("</tr>")     
        
        student_cnt = 0
        last_class = ""
        
        for student_idx in range(0, len(df.index)):
            student_cnt += 1
            stu_class = str(df.iloc[student_idx]['班級']).strip() 
            stu_name = str(df.iloc[student_idx]['姓名']).strip()
            
            # 分隔線邏輯與舊版一致
            if (last_class[0:1] != stu_class[0:1]) and (last_class != ""):
                f.write(break_line.format(7 + len(song_list) * 4))
            
            last_class = stu_class
            f.write("<tr>")
            
            html_student_str = html_student.replace("#STUDENT_NO#", str(student_cnt))
            stu_rank = str(df.iloc[student_idx]['總排名'])
            if stu_rank == "1":
                stu_rank = "🏆1"
            
            stu_score_tol = str(df.iloc[student_idx]['總積分'])
            stu_score_abs = str(df.iloc[student_idx]['出勤扣分'])
            stu_score_teach = str(df.iloc[student_idx]['老師加分'])
            
            html_student_str = html_student_str.replace("#STUDENT_CLASS#", stu_class)
            html_student_str = html_student_str.replace("#STUDENT_NAME#", stu_name)
            html_student_str = html_student_str.replace("#STUDENT_RANK#", stu_rank)
            
            if stu_score_abs != "0":
                html_student_str = html_student_str.replace("#ATTENT_SCORE#", stu_score_abs)
            else:    
                html_student_str = html_student_str.replace("#ATTENT_SCORE#", "")
             
            if stu_score_teach != "0":   
                html_student_str = html_student_str.replace("#TEACHER_SCORE#", stu_score_teach)
            else:
                html_student_str = html_student_str.replace("#TEACHER_SCORE#", "")
                
            if stu_score_tol == "None" or stu_score_tol == "":
                html_student_str = html_student_str.replace("#STUDENT_SCORE#", "NA")
            else:    
                html_student_st = str(stu_score_tol) + f"<progress class=\"a\" max={max_score} value={stu_score_tol}></progress>"
                html_student_str = html_student_str.replace("#STUDENT_SCORE#", html_student_st)
            
            if student_cnt % 2 == 0:
                html_student_str = html_student_str.replace("#BGCOLOR#", "#DCECEA")
            else:
                html_student_str = html_student_str.replace("#BGCOLOR#", "#FFFF2CC")
            f.write(html_student_str)
            
            for song_id, song_name in enumerate(song_list):
                html_student_rs_str = html_student_rs
                
                pass_score_val = str(df.iloc[student_idx][f'{song_id+1}_通過分數'])
                if pass_score_val == "0" or pass_score_val == "nan":
                    pass_score_val = "-"
                
                pass_time_val = str(df.iloc[student_idx][f'{song_id+1}_通過時間'])
                if pass_time_val == "nan":
                    pass_time_val = ""
                
                good_score_val = str(df.iloc[student_idx][f'{song_id+1}_很好加分'])
                if good_score_val == "0" or good_score_val == "nan":
                    good_score_val = ""
                
                demo_score_val = str(df.iloc[student_idx][f'{song_id+1}_示範加分'])
                if demo_score_val == "0" or demo_score_val == "nan":
                    demo_score_val = ""
                    
                html_student_rs_str = html_student_rs_str.replace("#上傳時間#", pass_time_val)
                html_student_rs_str = html_student_rs_str.replace("#優先通過分數#", pass_score_val)
                html_student_rs_str = html_student_rs_str.replace("#很好通過分數#", good_score_val)
                html_student_rs_str = html_student_rs_str.replace("#示範版分數#", demo_score_val)
                f.write(html_student_rs_str)
            f.write("</tr>")    
    print(f"- 已產生網頁檔案: {html_file_name}")

# ==============================================================================
# 3. 主自動化排程 (Main Pipeline)
# ==============================================================================
if __name__ == "__main__":
    config_file = 'ini/competition.ini'
    config = configparser.ConfigParser()
    
    with open(config_file, mode='rb') as f:
        content = f.read()
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    config.read_string(content.decode('utf8'))
    
    # 轉為小寫鍵值以匹配 INI 檔案
    pgm_info = {k.lower(): v for k, v in config['DEFAULT'].items()}
    
    grade_list = [g.strip() for g in pgm_info.get('grade', '').split(',') if g.strip()]
    song_list = [s.strip() for s in pgm_info.get('曲目資訊', '').split(',') if s.strip()]
    
    # --------------------------------------------------------------------------
    # 步驟 A：自動同步下載雲端 Excel 與建立歷史備份
    # --------------------------------------------------------------------------
    print("*下載最新資料檔案*" + "*"*30)
    adj_url = pgm_info.get('adj_url', '')
    if adj_url:
        DownLoadFile(adj_url, "加減分.xlsx")
        if check_md5("加減分.xlsx") != check_md5("加減分_BAK.xlsx"):
            shutil.copy("加減分.xlsx", "加減分_BAK.xlsx")
            config.set("DEFAULT", "adj_update", datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
            with open(config_file, 'w', encoding='utf-8') as cf:
                config.write(cf)
    
    # 重新讀取可能被更新的時間
    config.read(config_file, encoding='utf-8')
    pgm_info = {k.lower(): v for k, v in config['DEFAULT'].items()}
    pgm_info["adj_update_time"] = pgm_info.get("adj_update", "")
    
    update_string = ""
    for grade in grade_list:
        url_key = f"{grade}_url"
        grade_url = pgm_info.get(url_key, '')
        if grade_url:
            file_name = f"{grade}.xlsx"
            bak_name = f"{grade}_BAK.xlsx"
            DownLoadFile(grade_url, file_name)
            
            if check_md5(file_name) != check_md5(bak_name):
                shutil.copy(file_name, bak_name)
                config.set("DEFAULT", f"{grade}_update", datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
                with open(config_file, 'w', encoding='utf-8') as cf:
                    config.write(cf)
            
            # 重新抓取更新時間紀錄
            config.read(config_file, encoding='utf-8')
            current_info = {k.lower(): v for k, v in config['DEFAULT'].items()}
            grade_up_time = current_info.get(f"{grade}_update", "")
            if grade_up_time:
                update_string += f"{grade}更新時間:{grade_up_time} "
    
    pgm_info["_update_string"] = update_string

    # --------------------------------------------------------------------------
    # 步驟 B：多群組年級 Excel 的垂直橫向欄位完全載入
    # --------------------------------------------------------------------------
    print("*解析並合併多個年級 Excel*" + "*"*30)
    all_students_raw_data = []
    global_student_serial = 0
    
    for grade in grade_list:
        file_path = f"{grade}.xlsx"
        if not os.path.exists(file_path):
            file_path = f"{grade}_BAK.xlsx"
        if not os.path.exists(file_path):
            print(f" [警告] 找不到 {grade} 的成績檔案，跳過該群組。")
            continue
            
        df_grade = pd.read_excel(file_path)
        
        for idx in range(len(df_grade)):
            global_student_serial += 1
            stu_class = str(df_grade.iloc[idx, 1]).strip()
            stu_name = str(df_grade.iloc[idx, 2]).strip()
            
            student_row = [global_student_serial, stu_class, stu_name]
            
            for song_id, song_name in enumerate(song_list):
                pass_mode = ""
                pass_time = ""
                demo_score = 0
                good_score = 0
                pass_score = 0
                
                if song_name in df_grade.columns:
                    col_base_idx = list(df_grade.columns).index(song_name)
                    pass_mode = str(df_grade.iloc[idx, col_base_idx]).strip()
                    pass_time = str(df_grade.iloc[idx, col_base_idx + 1]).strip().replace("-", "/")
                    
                    if "★" in pass_mode:
                        demo_score = 5
                    if "●" in pass_mode:
                        good_score = 5
                        
                    if pass_mode in ("nan", "", "X", "□"):
                        pass_mode = ""
                        pass_time = ""
                
                student_row.extend([pass_mode, demo_score, good_score, pass_time, pass_score])
            
            all_students_raw_data.append(student_row)

    # 動態產生 DataFrame 的完整標準列欄位
    columns = ['序號', '班級', '姓名']
    for song_id, song_name in enumerate(song_list):
        columns.extend([
            f"{song_id+1}_通過方式", f"{song_id+1}_示範加分",
            f"{song_id+1}_很好加分", f"{song_id+1}_通過時間",
            f"{song_id+1}_通過分數"
        ])
    
    main_df = pd.DataFrame(all_students_raw_data, columns=columns)
    students_count = len(main_df.index)

    # --------------------------------------------------------------------------
    # 步驟 C：時間排序計分權重 (依通過時間給分，未通過者給 0 分)
    # --------------------------------------------------------------------------
    print("*核心計分排名運算中*" + "*"*30)
    for song_id, song_name in enumerate(song_list):
        sort_field = f'{song_id+1}_通過時間'
        score_field = f'{song_id+1}_通過分數'
        
        df_none_space = main_df[main_df[sort_field] != ''].copy()
        df_space = main_df[main_df[sort_field] == ''].copy()
        
        if len(df_none_space) > 0:
            df_none_space[score_field] = students_count - df_none_space[sort_field].rank(method='min').astype(int) + 1
            
        df_space[score_field] = 0
        main_df = pd.concat([df_none_space, df_space], ignore_index=True)

    # --------------------------------------------------------------------------
    # 步驟 D：不參賽名單過濾
    # --------------------------------------------------------------------------
    remove_students = pgm_info.get("remove_student", "").split(",")
    for stu_info in remove_students:
        if "_" in stu_info:
            stu_cls, stu_name = stu_info.split("_")
            main_df = main_df[~((main_df['班級'] == stu_cls.strip()) & (main_df['姓名'] == stu_name.strip()))]

    # --------------------------------------------------------------------------
    # 步驟 E：橫向對接「加減分.xlsx」調整檔 (复合主键 Left Join)
    # --------------------------------------------------------------------------
    adj_file = "加減分.xlsx" if os.path.exists("加減分.xlsx") else "加減分_BAK.xlsx"
    if os.path.exists(adj_file):
        df_adj = pd.read_excel(adj_file)
        df_adj['班級'] = df_adj['班級'].astype(str).str.strip()
        df_adj['姓名'] = df_adj['姓名'].astype(str).str.strip()
        df_adj = df_adj[['班級', '姓名', '出勤扣分', '老師加分']].copy()
        
        main_df = pd.merge(main_df, df_adj, on=['班級', '姓名'], how='left')
    else:
        main_df['出勤扣分'] = 0
        main_df['老師加分'] = 0

    main_df['出勤扣分'] = main_df['出勤扣分'].fillna(0).astype(int)
    main_df['老師加分'] = main_df['老師加分'].fillna(0).astype(int)

    # --------------------------------------------------------------------------
    # 步驟 F：點數大總結與總排名劃分
    # --------------------------------------------------------------------------
    sum_list = main_df['出勤扣分'] + main_df['老師加分']

    for song_id, song_name in enumerate(song_list):
        sum_list += main_df[f'{song_id+1}_通過分數'].astype(int)
        sum_list += main_df[f'{song_id+1}_示範加分'].astype(int)
        sum_list += main_df[f'{song_id+1}_很好加分'].astype(int)
        
    main_df['總積分'] = sum_list
    main_df['總排名'] = main_df['總積分'].rank(method='min', ascending=False).astype(int)

    # 寫入 Result.xlsx 供人工核對備份
    try:
        main_df.to_excel("Result.xlsx", index=False)
    except:
        pass

    # --------------------------------------------------------------------------
    # 步驟 G：排序並輸出為與原格式相同的 HTML
    # --------------------------------------------------------------------------
    df_final = main_df.sort_values('序號')
    GenReport(df_final, song_list, pgm_info)
    print("\n====================================================")
    print(" 本地成績處理與網頁 HTML 生成已順利執行完畢。")
    print("====================================================")

    # --------------------------------------------------------------------------
    # 擴充功能 H：讀取 INI 設定並透過 Access Token 自動部署至 GitHub
    # --------------------------------------------------------------------------
    repo_url = pgm_info.get('github_url', '').strip()
    token = pgm_info.get('github_token', '').strip()

    if repo_url and token:
        print("*啟動 GitHub Pages 雲端自動同步管線*" + "*"*20)
        import subprocess
        
        # 1. 如果尚未建立本地 Git 儲存庫，自動進行初始化
        if not os.path.exists(".git"):
            print(" [資訊] 偵測到本地儲存庫尚未初始化，正在自動建立連結...")
            subprocess.run(["git", "init"], capture_output=True)
            subprocess.run(["git", "branch", "-M", "main"], capture_output=True)

        # 2. 清洗網址並將 Token 強制內嵌至 URL 中以繞過密碼認證
        # 將 https://github.com/... 轉換為 https://<token>@github.com/...
        clean_url = repo_url.replace("https://", "").replace("http://", "")
        authenticated_url = f"https://{token}@{clean_url}.git"

        # 3. 重新設定遠端倉庫網址 (確保每次都使用最新 Token)
        subprocess.run(["git", "remote", "remove", "origin"], capture_output=True)
        subprocess.run(["git", "remote", "add", "origin", authenticated_url], capture_output=True)

        # 4. 執行標準的 Git 提交管線
        print(" [Git] 正在封裝網頁變更元件...")
        subprocess.run(["git", "add", "."])
        
        commit_msg = f"Auto Update: Score HTML Report ({datetime.now().strftime('%Y/%m/%d %H:%M:%S')})"
        subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True)

        print(" [Git] 正在透過安全通道推送至 GitHub Pages (main)...")
        # 追蹤推送狀態
        push_result = subprocess.run(["git", "push", "-u", "origin", "main", "--force"], capture_output=True, text=True)
        
        if push_result.returncode == 0:
            print("🏆 [成功] 網頁已透過 Token 完美部署至雲端 GitHub Pages！")
        else:
            print("❌ [失敗] Git 推送失敗。錯誤訊息如下：")
            print(push_result.stderr)
    else:
        print("\n[提示] INI 設定檔中未偵測到 github_url 或 github_token，跳過雲端自動同步。")

    print("\n====================================================")
    print(" 全自動管線計算與網頁更新程序已全部順利結束。")
    print("====================================================")
