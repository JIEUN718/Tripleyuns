import win32com.client
import datetime
import os


def download_and_merge_today_ppts():
    # --- [0단계] 부서 및 담당자 정보 (취합 순서 기준) ---
    departments = [
        {"order": 1, "team": "연구기획팀", "sender": "허용민"},
        {"order": 2, "team": "심혈관팀", "sender": "김태원"},
        {"order": 3, "team": "급성감염팀", "sender": "한예지"},
        {"order": 4, "team": "Cancer팀", "sender": "정진용"},
        {"order": 5, "team": "호르몬팀", "sender": "이소희"},
        {"order": 6, "team": "치료용항체팀", "sender": "김영은"},
        {"order": 7, "team": "갑상선팀", "sender": "김세희"},
        {"order": 8, "team": "당뇨팀", "sender": "함은선"}
    ]

    submitted_teams = set()

    # --- [1단계] 저장할 폴더 설정 ---
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    save_dir = os.path.join(desktop_path, "오늘의_회의자료")

    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    # --- [2단계] 아웃룩 메일 검색 및 다운로드 ---
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    target_folder = None
    account_email = "jje@boditech.co.kr"

    for folder in outlook.Folders:
        if account_email in folder.Name:
            try:
                target_folder = folder.Folders("받은 편지함").Folders("내 메일")
                break
            except Exception:
                pass

    if target_folder is None:
        try:
            target_folder = outlook.GetDefaultFolder(6).Folders("내 메일")
        except Exception:
            print("❌ '내 메일' 폴더를 찾을 수 없습니다.")
            return

    items = target_folder.Items
    items.Sort("[ReceivedTime]", True)

    today = datetime.date.today()
    today_str = today.strftime("%y%m%d")

    keywords = ['주간회의자료', '회의자료']
    ppt_exts = ['.ppt', '.pptx']

    downloaded_files = []
    count = 0

    print(f"🔍 아웃룩 탐색 시작 (조건: {today_str} 날짜가 포함된 PPT 파일)...\n")

    for item in items:
        try:
            if item.Class != 43:
                continue

            received_date = item.ReceivedTime.date()
            if received_date < today:
                break
            if received_date != today:
                continue

            subject = item.Subject if item.Subject else ""
            body = item.Body if item.Body else ""
            sender_name = item.SenderName

            if item.Attachments.Count > 0:
                for att in item.Attachments:
                    fname = att.FileName.lower()

                    if any(fname.endswith(ext) for ext in ppt_exts):
                        if today_str in fname:
                            if (any(kw in att.FileName for kw in keywords) or
                                    any(kw in subject for kw in keywords) or
                                    any(kw in body for kw in keywords)):

                                count += 1
                                safe_filename = f"{count}_{att.FileName}"
                                save_path = os.path.abspath(os.path.join(save_dir, safe_filename))

                                att.SaveAsFile(save_path)

                                matched_order = 99
                                matched_team = "기타(명단 외)"

                                for dept in departments:
                                    if dept["sender"] in sender_name:
                                        submitted_teams.add(dept["team"])
                                        matched_order = dept["order"]
                                        matched_team = dept["team"]
                                        break

                                downloaded_files.append((matched_order, save_path))
                                print(f"[{count}] 다운로드 완료: {safe_filename} - {matched_team}({sender_name})")

        except Exception as e:
            continue

    # --- [2-1단계] 제출 현황 및 미제출 부서 출력 ---
    print("\n==================================")
    print("      [주간보고서 제출 현황]")
    print("==================================")

    missing_teams = [dept for dept in departments if dept["team"] not in submitted_teams]

    if missing_teams:
        print("❌ 미제출 부서 목록:")
        for dept in missing_teams:
            print(f"  - {dept['order']}번. {dept['team']} (담당자: {dept['sender']})")
    else:
        print("✅ 모든 부서가 주간보고서를 제출했습니다.")
    print("==================================\n")

    # --- [3단계] 파워포인트 파일 병합 ---
    if not downloaded_files:
        print("조건에 맞는(날짜 포함) PPT 파일이 없습니다. 프로그램을 종료합니다.")
        return

    print("🔄 파워포인트를 실행하여 파일 병합을 시작합니다...")

    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        merged_prs = ppt_app.Presentations.Add(WithWindow=False)

        # 다운로드된 파일들을 '자료취합순서' 기준으로 정렬
        downloaded_files.sort(key=lambda x: x[0])

        for order, file_path in downloaded_files:
            current_slide_count = merged_prs.Slides.Count
            merged_prs.Slides.InsertFromFile(file_path, current_slide_count)
            print(f" - 병합됨 (순서 {order if order != 99 else '미상'}): {os.path.basename(file_path)}")

        # 💡 [핵심 해결] WinError 32 방지 로직
        final_file_name = "주간보고병합.pptx"
        final_save_path = os.path.abspath(os.path.join(save_dir, final_file_name))

        # 파일이 이미 존재하면 삭제 시도
        if os.path.exists(final_save_path):
            try:
                os.remove(final_save_path)
            except PermissionError:
                # 삭제 불가능(열려 있거나 락 걸림)할 경우 시간값을 뒤에 붙여 새 이름으로 저장
                now_str = datetime.datetime.now().strftime("%H%M%S")
                final_file_name = f"주간보고병합_{now_str}.pptx"
                final_save_path = os.path.abspath(os.path.join(save_dir, final_file_name))
                print(f"\n⚠️ [경고] 기존 '주간보고병합.pptx' 파일이 열려 있어 덮어쓸 수 없습니다.")
                print(f"👉 대안으로 새 파일명 '{final_file_name}'으로 저장합니다.")

        merged_prs.SaveAs(final_save_path)
        merged_prs.Close()
        ppt_app.Quit()

        print(f"\n✅ [성공] 모든 파일이 성공적으로 병합되었습니다!")
        print(f"📁 저장 위치: {final_save_path}")

    except Exception as e:
        print(f"❌ 파워포인트 병합 중 다른 오류가 발생했습니다: {e}")
        try:
            ppt_app.Quit()
        except:
            pass


if __name__ == "__main__":
    download_and_merge_today_ppts()
