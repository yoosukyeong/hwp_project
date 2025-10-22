# 해설 빈 파일 찾는 코드

# -*- coding: utf-8 -*-
import os
import win32com.client
import win32clipboard
import traceback

# ================= 사용자 설정 =================
INPUT_HWP_ROOT   = r"C:\Users\tnwls\OneDrive\바탕 화면\서체변환_한글파일\hwp_organized_20250908\1"   # 루트 폴더 (하위폴더 전부 탐색)
INCLUDE_HWPX     = False                       # .hwpx도 포함하려면 True
REQUIRE_ANY_TABLE_FOR_CONTENT = False          # 표가 없으면 '내용없음'으로 간주하려면 True
LOG_FILE         = r"C:\Users\tnwls\OneDrive\바탕 화면\서체변환_한글파일\hwp_organized_20250908\해설빈파일_log.txt"
# =================================================

def create_hwp():
    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass
    try:
        hwp.SetMessageBoxMode(0)
    except Exception:
        pass
    return hwp

def get_plain_text(hwp) -> str:
    try:
        t = hwp.GetTextFile("TEXT", "")
        return (t or "").replace("\r", "").replace("\n", "").strip()
    except:
        return ""

def get_text_via_clipboard(hwp) -> str:
    try:
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Copy")
        win32clipboard.OpenClipboard()
        try:
            import win32con
            data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        except Exception:
            data = ""
        finally:
            win32clipboard.CloseClipboard()
        hwp.HAction.Run("Cancel")
        return (data or "").replace("\r", "").replace("\n", "").strip()
    except:
        try:
            hwp.HAction.Run("Cancel")
        except:
            pass
        return ""

def count_controls(hwp) -> dict:
    counts = {"tbl":0, "gso":0, "eqed":0, "pic":0, "ole":0}
    try:
        ctrl = hwp.HeadCtrl
        while ctrl:
            cid = (getattr(ctrl, "CtrlID", "") or "").lower()
            if cid in counts:
                counts[cid] += 1
            ctrl = ctrl.Next
    except Exception:
        pass
    return counts

def has_any_content(hwp) -> bool:
    c = count_controls(hwp)

    if REQUIRE_ANY_TABLE_FOR_CONTENT and c["tbl"] == 0:
        return False

    text1 = get_plain_text(hwp)
    if text1:
        return True

    text2 = get_text_via_clipboard(hwp)
    if text2:
        return True

    if (c["tbl"] + c["gso"] + c["eqed"] + c["pic"] + c["ole"]) > 0:
        return True

    return False

def is_target_file(fname: str) -> bool:
    f = fname.lower()
    if f.startswith("~$"):
        return False
    if f.endswith(".hwp"):
        return True
    if INCLUDE_HWPX and f.endswith(".hwpx"):
        return True
    return False

def iter_all_files(root_dir: str):
    for r, _, files in os.walk(root_dir):
        for f in files:
            if is_target_file(f):
                yield os.path.join(r, f)

def main():
    hwp = create_hwp()
    total = 0
    empty = 0
    error = 0

    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write("\n\n==== 검사 시작 ====\n")
        for path in iter_all_files(INPUT_HWP_ROOT):
            total += 1
            print(f"[{total}] 검사 중: {path}")
            try:
                hwp.Open(path)
                if not has_any_content(hwp):
                    msg = f"{path} 내용없음"
                    log.write(msg + "\n")
                    print("  → 내용없음")
                    empty += 1
                else:
                    print("  → 내용 있음")
            except Exception as e:
                msg = f"{path} 오류: {repr(e)}"
                log.write(msg + "\n")
                print("  → 오류:", repr(e))
                error += 1
            finally:
                try:
                    hwp.Clear(3)
                except:
                    pass

        summary = f"검사 완료: 총 {total}개 / 내용없음 {empty}개 / 오류 {error}개\n"
        log.write(summary)
        print(summary)


if __name__ == "__main__":
    main()
