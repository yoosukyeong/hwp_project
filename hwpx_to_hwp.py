# coding: utf-8
import os
import time
import pythoncom
import win32com.client as win32
import json

# -------------------------
# HWP 제어
# -------------------------
def create_hwp(hide_window: bool = True):
    created_app = False
    try:
        hwp = win32.GetActiveObject('HWPFrame.HwpObject')
    except Exception:
        hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
        created_app = True

    # 보안 모듈(경로체크) 우회
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass

    try:
        hwp.SetMessageBoxMode(0)  # 0: 안보임, 1: 보임
    except Exception:
        pass

    if hide_window:
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass

    return hwp, created_app


def open_hwp(hwp, target_path: str) -> bool:

    if not os.path.exists(target_path):
        return False
    open_arg = "versionwarning:false;lock:false"
    try:
        return bool(hwp.Open(target_path, arg=open_arg))
    except Exception:
        return False


def close_hwp(hwp):
    try:
        hwp.HAction.GetDefault("FileClose", hwp.HParameterSet.HFileClose.HSet)
        hwp.HParameterSet.HFileClose.IsSave = 0
        hwp.HAction.Execute("FileClose", hwp.HParameterSet.HFileClose.HSet)
    except Exception:
        try:
            hwp.Clear(3)
        except Exception:
            pass

# -------------------------
# HWPX → HWP 변환
# -------------------------
def hwpx_to_hwp(hwp, hwpx_path: str, out_path: str = None) -> str:
    if not os.path.exists(hwpx_path):
        return ""

    base_dir = os.path.dirname(hwpx_path)
    base_name = os.path.splitext(os.path.basename(hwpx_path))[0]
    if out_path is None:
        out_path = os.path.join(base_dir, base_name + ".hwp")

    # 원본 열기
    if not open_hwp(hwp, hwpx_path):
        return ""

    try:
        # 액션 기반 Save As
        try:
            hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileSaveAs.HSet)
            hwp.HParameterSet.HFileSaveAs.FileName = out_path
            # 포맷 지정
            try:
                hwp.HParameterSet.HFileSaveAs.Format = "HWP"
            except Exception:
                try:
                    hwp.HParameterSet.HFileSaveAs.FormatShortName = "HWP"
                except Exception:
                    pass
            hwp.HParameterSet.HFileSaveAs.OverWrite = 1
            hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileSaveAs.HSet)
        except Exception:
            # SaveAs 직접 호출
            try:
                hwp.SaveAs(out_path, "HWP")
            except Exception:
                return ""

        return out_path if os.path.exists(out_path) else ""
    finally:
        close_hwp(hwp)



# -------------------------
# 단일 파일 변환 진입점 (C# or 직접 호출)
# -------------------------
def convert_hwpx_to_hwp(file_path: str) -> str:
    """
    C# 또는 Python에서 직접 호출 가능.
    예: result = convert_hwpx_to_hwp(r"D:\\problem_ex\\test\\A_001.hwpx")
    """
    pythoncom.CoInitialize()
    hwp, created = create_hwp(hide_window=True)
    try:
        out_path = hwpx_to_hwp(hwp, file_path)
        return out_path or ""
    finally:
        try:
            if created:
                hwp.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# -------------------------
# 실행부 (폴더 내 모든 hwpx 변환)
# -------------------------
# if __name__ == "__main__":
##     c# 에서 아래처럼 호출
#     test_file = r"D:\problem_ex\test\A9208B0171_p.hwpx"
#     result = convert_hwpx_to_hwp(test_file)


#     if result:
#         print("✅ 변환 완료:", result)
#     else:
#         print("❌ 변환 실패:", test_file)
