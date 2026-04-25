from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import re

# =========================
# 1. 설정 변수
# =========================
TARGET_GRADE = 3  # 작업할 학년 (2 또는 3)
START_COURSE_INDEX = 1  # ★ 작업 시작할 개설과목 순번 (1부터 시작). 예: 4입력 시 4번째 과목부터 이어하기
FILE_ALL_STUDENTS = "★☆2026학년도 1학기 주문형 교육과정 전체 학생 명단☆★.xlsx"
FILE_CLASS_ASSIGN = "★☆2026학년도 1학기 주문형 교육과정 반 편성 자료☆★.xlsx"

# =========================
# 2. 헬퍼 함수
# =========================
def print_log(msg):
    print(msg, flush=True)

def normalize_text(text):
    if not text or pd.isna(text): return ""
    text = str(text).strip().upper()
    text = text.replace("II", "Ⅱ").replace("I", "Ⅰ")
    text = re.sub(r'\s+', '', text)
    return text

def clean_num(val):
    """엑셀에서 번호나 반이 '1.0' 처럼 소수점으로 나오는 현상 방지"""
    try:
        return str(int(float(val)))
    except:
        return str(val).strip()

def sort_class_key(cls_str):
    """반 번호를 오름차순(1반->2반->3반)으로 정렬하기 위한 키 함수"""
    try:
        return int(re.sub(r'[^0-9]', '', cls_str))
    except:
        return 999

# =========================
# 3. 데이터 로드 (Pandas) - [STEP 2, 3, 4] 사전 매핑
# =========================
def prepare_data():
    print_log("📊 엑셀 데이터를 불러오고 매핑하는 중...")
    df_all = pd.read_excel(FILE_ALL_STUDENTS, sheet_name="1학기")
    df_assign = pd.read_excel(FILE_CLASS_ASSIGN, sheet_name=f"{TARGET_GRADE}학년")

    assign_map = {}
    # [파일 2]에서 고유 키 조합 -> (편성 반, 편성 번호) 매핑 사전 생성
    for _, row in df_assign.iterrows():
        sch = str(row.iloc[5]).strip()       
        grd = clean_num(row.iloc[7])         
        org_cls = clean_num(row.iloc[9])     
        org_num = clean_num(row.iloc[10])    
        name = str(row.iloc[11]).strip()     
        
        assigned_cls = clean_num(row.iloc[1]) 
        assigned_num = clean_num(row.iloc[2]) 
        
        key = f"{sch}_{grd}_{org_cls}_{org_num}_{name}"
        assign_map[key] = (assigned_cls, assigned_num)

    course_dict = {}
    for _, row in df_all.iterrows():
        is_registered = str(row.iloc[10]).strip().upper() 
        if is_registered == 'FALSE' or is_registered == 'F':
            continue
            
        course_name = normalize_text(row.iloc[11]) 
        if not course_name or course_name == 'NAN':
            continue
            
        sch = str(row.iloc[3]).strip()       
        grd = clean_num(row.iloc[4])         
        org_cls = clean_num(row.iloc[5])     
        org_num = clean_num(row.iloc[6])     
        name = str(row.iloc[7]).strip()      
        
        key = f"{sch}_{grd}_{org_cls}_{org_num}_{name}"
        matched_cls, matched_num = None, None
        
        if key in assign_map:
            matched_cls, matched_num = assign_map[key]
        else:
            matches = [v for k, v in assign_map.items() if k.startswith(f"{sch}_") and k.endswith(f"_{name}")]
            if len(matches) == 1: 
                matched_cls, matched_num = matches[0]

        if matched_cls is not None and matched_num is not None:
            if course_name not in course_dict:
                course_dict[course_name] = []
            course_dict[course_name].append({
                'cls': matched_cls,
                'num': matched_num,
                'name': name
            })
            
    print_log("✅ 엑셀 데이터 매핑 완료!\n")
    return course_dict

# =========================
# 4. 나이스 UI 제어 함수
# =========================
def switch_to_neis_window(driver):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        time.sleep(0.3)
        if "neis.go.kr" in driver.current_url:
            return True
    return False

def get_total_course_count(driver):
    return driver.execute_script("""
        let grid = document.querySelector('div[aria-label="개설과목"][role="grid"]');
        if(grid) {
            let rowCount = parseInt(grid.getAttribute('aria-rowcount'));
            return isNaN(rowCount) ? 0 : rowCount - 1; 
        }
        return 0;
    """)

def click_course_by_index(driver, index):
    for _ in range(20):
        script = f"""
        let grid = document.querySelector('div[aria-label="개설과목"][role="grid"]');
        if(!grid) return 'NO_GRID';
        
        // 처음 시작하거나, 위로 강제 이동이 필요한 경우 스크롤 리셋
        if ({index} === 0 || !window.resetScrollDone) {{
            let scrollbar = grid.querySelector('.cl-scrollbar');
            if(scrollbar) {{ scrollbar.scrollTop = 0; scrollbar.dispatchEvent(new Event('scroll', {{bubbles: true}})); }}
            window.resetScrollDone = true;
        }}

        let rows = grid.querySelectorAll('.cl-grid-row');
        for(let row of rows) {{
            let rIdx = row.getAttribute('data-rowindex');
            if(rIdx == '{index}') {{
                let cell = row.querySelector('div[data-cellindex="2"] .cl-text');
                if(cell) {{
                    let courseName = cell.innerText || cell.textContent;
                    cell.dispatchEvent(new MouseEvent('mousedown', {{bubbles: true}}));
                    cell.dispatchEvent(new MouseEvent('mouseup', {{bubbles: true}}));
                    cell.click();
                    return courseName.trim();
                }}
            }}
        }}
        let scrollbar = grid.querySelector('.cl-scrollbar');
        if(scrollbar) {{
            scrollbar.scrollTop += 60; 
            scrollbar.dispatchEvent(new Event('scroll', {{bubbles: true}}));
        }}
        return 'NOT_FOUND';
        """
        res = driver.execute_script(script)
        if res not in ['NO_GRID', 'NOT_FOUND']:
            return res
        time.sleep(0.3)
    return None

def change_combobox(driver, box_type, target_text):
    if box_type == "교생": label_keyword = "교생"
    elif box_type == "학년": label_keyword = "학년"
    elif box_type == "반": label_keyword = "반"
        
    # 1. 콤보박스 클릭 (목록 열기)
    driver.execute_script(f"""
        let combos = Array.from(document.querySelectorAll('div[role="combobox"]'));
        let targetCombos = combos.filter(c => c.getAttribute('aria-label') && c.getAttribute('aria-label').includes('{label_keyword}'));
        if(targetCombos.length > 0) {{
            let combo = targetCombos[targetCombos.length - 1];
            combo.dispatchEvent(new MouseEvent('mousedown', {{bubbles: true}}));
            combo.dispatchEvent(new MouseEvent('mouseup', {{bubbles: true}}));
            combo.click();
        }}
    """)
    time.sleep(0.5)
    
    # 2. 팝업 스크롤 맨 위로 초기화 (중요: 이전 위치 기억 리셋)
    driver.execute_script("""
        let popups = Array.from(document.querySelectorAll('.cl-popup'));
        let activePopup = popups[popups.length - 1];
        if(activePopup) {
            let scrollbar = activePopup.querySelector('.cl-scrollbar') || activePopup;
            if(scrollbar) {
                scrollbar.scrollTop = 0;
                scrollbar.dispatchEvent(new Event('scroll', {bubbles: true}));
            }
        }
    """)
    time.sleep(0.3)

    # 3. 찾을 때까지 스크롤 내리며 클릭
    for _ in range(15):
        script = f"""
        let popups = Array.from(document.querySelectorAll('.cl-popup'));
        let activePopup = popups[popups.length - 1];
        if(!activePopup) return 'NO_POPUP';
        
        let items = Array.from(activePopup.querySelectorAll('.cl-text'));
        let matches = items.filter(el => el.innerText.trim() === '{target_text}' && el.offsetWidth > 0);
        
        if(matches.length > 0) {{
            let targetItem;
            // 1반일 경우에만 맨 위 가짜 1반을 피하기 위해 배열의 마지막(진짜 1반)을 선택
            if ('{target_text}' === '1반' && matches.length > 1) {{
                targetItem = matches[matches.length - 1];
            }} else {{
                targetItem = matches[0];
            }}
            
            targetItem.scrollIntoView({{block: 'center'}});
            targetItem.dispatchEvent(new MouseEvent('mousedown', {{bubbles: true}}));
            targetItem.dispatchEvent(new MouseEvent('mouseup', {{bubbles: true}}));
            targetItem.click();
            return 'FOUND';
        }}
        
        // 못 찾았으면 스크롤 내리기
        let scrollbar = activePopup.querySelector('.cl-scrollbar') || activePopup;
        let oldTop = scrollbar.scrollTop;
        scrollbar.scrollTop += 100;
        scrollbar.dispatchEvent(new Event('scroll', {{bubbles: true}}));
        if(oldTop === scrollbar.scrollTop) return 'BOTTOM'; // 바닥 도달
        return 'NOT_FOUND';
        """
        res = driver.execute_script(script)
        if res == 'FOUND':
            time.sleep(0.5)
            return
        elif res == 'BOTTOM' or res == 'NO_POPUP':
            break
        time.sleep(0.2)
        
    # 못 찾았을 경우 팝업 닫기 처리
    driver.execute_script("document.body.click();")
    time.sleep(0.5)

def click_action_btn(driver, btn_label):
    driver.execute_script(f"""
        let btns = Array.from(document.querySelectorAll('div[aria-label="{btn_label}"][role="button"]'));
        if (btns.length > 0) {{
            btns[btns.length - 1].click();
        }}
    """)
    time.sleep(1.0)

def check_students_in_bulk(driver, students_list):
    targets = [{'num': re.sub(r'[^0-9]', '', str(s['num'])), 'name': s['name'].replace(' ', '')} for s in students_list]
    found_count = 0
    
    driver.execute_script("""
        let grid = document.querySelector('div[aria-label="미편성학생"][role="grid"]');
        if(grid) {
            let scrollbar = grid.querySelector('.cl-scrollbar');
            if(scrollbar) { scrollbar.scrollTop = 0; scrollbar.dispatchEvent(new Event('scroll', {bubbles: true})); }
        }
    """)
    time.sleep(0.5)

    for _ in range(30): 
        if not targets:
            break 
            
        found_in_view = driver.execute_script("""
            let targets = arguments[0];
            let grid = document.querySelector('div[aria-label="미편성학생"][role="grid"]');
            let found = [];
            if(!grid) return found;
            
            let rows = grid.querySelectorAll('.cl-grid-row');
            for(let row of rows) {
                let numCell = row.querySelector('div[data-cellindex="3"] .cl-text');
                let nameCell = row.querySelector('div[data-cellindex="4"] .cl-text');
                let checkbox = row.querySelector('.cl-checkbox-icon');
                
                if(numCell && nameCell && checkbox) {
                    let cNum = numCell.innerText.replace(/[^0-9]/g, ''); 
                    let cName = nameCell.innerText.replace(/\\s+/g, '');
                    
                    let matchIdx = targets.findIndex(t => t.num === cNum && t.name === cName);
                    if(matchIdx !== -1) {
                        let parentBox = checkbox.closest('.cl-checkbox');
                        let isChecked = (checkbox.getAttribute('aria-checked') === 'true') || 
                                        (parentBox && parentBox.classList.contains('cl-checked'));
                        
                        found.push({
                            element: checkbox,
                            num: targets[matchIdx].num,
                            name: targets[matchIdx].name,
                            needsClick: !isChecked
                        });
                        targets.splice(matchIdx, 1);
                    }
                }
            }
            return found;
        """, targets)
        
        if found_in_view:
            for f in found_in_view:
                if f['needsClick']:
                    driver.execute_script("arguments[0].click();", f['element'])
                    time.sleep(0.15) 
                    
                print_log(f"      ✔️ [체크 완료] {f['num']}번 {f['name']}")
                found_count += 1
                targets = [t for t in targets if not (t['num'] == f['num'] and t['name'] == f['name'])]
                
        if not targets:
            break
            
        at_bottom = driver.execute_script("""
            let grid = document.querySelector('div[aria-label="미편성학생"][role="grid"]');
            let scrollbar = grid.querySelector('.cl-scrollbar');
            if(!scrollbar) return true;
            let oldTop = scrollbar.scrollTop;
            scrollbar.scrollTop += 150;
            scrollbar.dispatchEvent(new Event('scroll', {bubbles: true}));
            return oldTop === scrollbar.scrollTop; 
        """)
        
        if at_bottom:
            break
        time.sleep(0.4)
        
    return found_count, targets

def handle_alerts_html(driver):
    def click_modal_confirm():
        script = """
        let dialogs = Array.from(document.querySelectorAll('.cl-dialog'));
        if (dialogs.length > 0) {
            let topDialog = dialogs[dialogs.length - 1];
            let btns = Array.from(topDialog.querySelectorAll('.cl-button'));
            let confirmBtn = btns.find(btn => btn.innerText.includes('확인'));
            
            if (confirmBtn) {
                confirmBtn.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                confirmBtn.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
                confirmBtn.click();
                return true;
            }
        }
        return false;
        """
        return driver.execute_script(script)

    clicked_first = False
    for _ in range(10): 
        time.sleep(0.5)
        if click_modal_confirm():
            clicked_first = True
            break
    if not clicked_first:
        print_log("  ⚠️ '저장하시겠습니까?' 팝업을 찾지 못했습니다.")

    clicked_second = False
    for _ in range(30): 
        time.sleep(0.5)
        if click_modal_confirm():
            clicked_second = True
            break
    if not clicked_second:
        print_log("  ⚠️ '저장되었습니다' 팝업을 찾지 못했습니다. 수동 확인이 필요할 수 있습니다.")
    time.sleep(1.0) 

# =========================
# 5. 메인 실행 루프
# =========================
def run():
    print_log(f"🚀 나이스 수강생 편성 자동화 ({TARGET_GRADE}학년) 시작\n")
    course_data = prepare_data()
    
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    if not switch_to_neis_window(driver): 
        print_log("❌ 나이스 창을 찾을 수 없습니다.")
        return

    driver.execute_script("window.resetScrollDone = false;")
    total_courses = get_total_course_count(driver)
    print_log(f"📌 나이스 화면에서 총 {total_courses}개의 개설과목을 감지했습니다.\n")
    
    summary_report = []

    # === 화면 순서대로 순회 시작 ===
    for i in range(total_courses):
        
        # ★ 특정 시작 번호부터 이어하기 로직
        if i < START_COURSE_INDEX - 1:
            print_log(f"⏭️ [{i+1}/{total_courses}] 스킵됨 (START_COURSE_INDEX 설정)")
            continue
            
        ui_course_name = click_course_by_index(driver, i)
        
        if not ui_course_name:
            print_log(f"⚠️ {i+1}번째 항목을 화면에서 찾을 수 없어 작업을 중단합니다.")
            break
            
        print_log("-" * 65)
        print_log(f"▶️ [{i+1}/{total_courses}] {ui_course_name} 작업 시작...")
        
        norm_name = normalize_text(ui_course_name)
        if norm_name not in course_data:
            print_log(f"  ⏭️ 해당 강의실은 엑셀 명단에 없거나 모두 'FALSE' 처리되어 스킵합니다.")
            summary_report.append({"course": ui_course_name, "assigned": 0, "total": 0, "failed_logs": ["명단에 없음/FALSE"]})
            continue
            
        students = course_data[norm_name]
        print_log(f"  👥 편성 대상: 총 {len(students)}명")
        
        # [추가] 강의실 클릭 후 우측 미편성학생 UI가 갱신될 때까지 안전하게 대기
        print_log("  ⏳ UI 화면 갱신 대기 중...")
        time.sleep(2.0) 
        
        change_combobox(driver, "교생", "타교생")
        change_combobox(driver, "학년", f"{TARGET_GRADE}학년")
        
        assigned_count = 0
        failed_logs = []
        
        # 반 목록을 추출한 뒤, 숫자로 오름차순(1반->2반->3반) 정렬하여 작업 최적화
        classes_to_process = sorted(list(set([s['cls'] for s in students])), key=sort_class_key)
        
        for cls in classes_to_process:
            change_combobox(driver, "반", f"{cls}반")
            click_action_btn(driver, "조회")
            time.sleep(2.0)
            
            cls_students = [s for s in students if s['cls'] == cls]
            found_count, missing_students = check_students_in_bulk(driver, cls_students)
            assigned_count += found_count
            
            for m in missing_students:
                failed_logs.append(f"{cls}반 {m['num']}번 {m['name']}")
                    
            if found_count > 0:
                click_action_btn(driver, "추가")
                time.sleep(1.0)
            
        if assigned_count > 0:
            click_action_btn(driver, "저장")
            handle_alerts_html(driver)
        
        print_log(f"  🎉 완료: {assigned_count}/{len(students)}명 성공")
        
        if failed_logs:
            print_log(f"  ⚠️ 실패 내역: {', '.join(failed_logs)}")

        summary_report.append({
            "course": ui_course_name,
            "assigned": assigned_count,
            "total": len(students),
            "failed_logs": failed_logs
        })

    print_log("\n" + "="*70)
    print_log(f"🏁 화면에 있는 모든 개설과목 작업이 끝났습니다.")
    print_log("-" * 70)
    print_log("📊 [최종 요약 보고서]")
    print_log("-" * 70)
    for report in summary_report:
        if report['total'] == 0:
            print_log(f"➖ {report['course']} : 명단에 없음 스킵됨")
        else:
            status = "✅" if report['assigned'] == report['total'] else "⚠️"
            print_log(f"{status} {report['course']} : {report['assigned']}/{report['total']}명 편성 완료")
            if report['failed_logs']:
                print_log(f"    └ 누락/실패: {', '.join(report['failed_logs'])}")
    print_log("="*70 + "\n")

if __name__ == "__main__":
    run()