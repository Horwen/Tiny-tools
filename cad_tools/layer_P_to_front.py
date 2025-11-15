# -*- coding: utf-8 -*-
"""
批量将当前目录下所有 DWG 文件中图层 P 的对象置于最前（DRAWORDER Front）
用法：
    1) 安装 pywin32:  pip install pywin32
    2) 将本脚本放到 DWG 所在目录
    3) 在该目录运行:  python layer_P_to_front.py
"""

import os
import glob
import time

import win32com.client
import pythoncom

TARGET_LAYER_NAME = "P"


def get_autocad_app():
    """
    连接/启动 AutoCAD，优先复用已打开实例。
    如需要可以把 ProgID 改成自己机器上实际的。
    """
    # 先尝试复用已打开的 AutoCAD
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("已复用正在运行的 AutoCAD 实例。")
        return acad
    except Exception:
        pass

    # 启动新的 AutoCAD 实例
    # 如报错，可尝试改成 "AutoCAD.Application.24" 或 "AutoCAD.Application.2024"
    prog_ids = [
        "AutoCAD.Application",
        "AutoCAD.Application.24",
        "AutoCAD.Application.2024",
    ]

    last_err = None
    for pid in prog_ids:
        try:
            print(f"尝试通过 ProgID: {pid} 启动 AutoCAD ...")
            acad = win32com.client.Dispatch(pid)
            print(f"成功通过 {pid} 启动 AutoCAD。")
            return acad
        except Exception as e:
            last_err = e
            print(f"  失败：{e!r}")

    raise RuntimeError(f"无法连接 AutoCAD，最后错误：{last_err!r}")


def main():
    pythoncom.CoInitialize()

    cwd = os.getcwd()
    dwg_files = glob.glob(os.path.join(cwd, "*.dwg"))

    if not dwg_files:
        print("当前目录中未找到任何 DWG 文件。")
        return

    print("将在当前目录处理以下 DWG 文件：")
    for f in dwg_files:
        print("  -", os.path.basename(f))

    # 连接 AutoCAD
    try:
        acad = get_autocad_app()
    except Exception as e:
        print("无法连接 AutoCAD，请检查：")
        print("  1）AutoCAD 2024 是否安装并能正常打开；")
        print("  2）Python 是否为 64 位；")
        print("  3）是否有权限问题，可尝试管理员运行；")
        print("  4）如仍不行，把实际 ProgID 告诉我再调。")
        print("详细错误信息：", e)
        return

    # 需要可改为 False
    acad.Visible = True

    # AutoLISP 命令：把图层 P 上的对象置于最前
    # 逻辑：
    #   (if (setq ss (ssget "X" '((8 . "P"))))
    #       (command "_.DRAWORDER" ss "" "F")
    #   )
    lisp_cmd_template = (
        '(progn '
        '(if (setq ss (ssget "X" \'((8 . "{layer}")))) '
        '  (command "_.DRAWORDER" ss "" "F" )'
        ')'
        ')\n'
    )
    lisp_cmd = lisp_cmd_template.format(layer=TARGET_LAYER_NAME)

    for dwg_path in dwg_files:
        print("\n处理文件：", os.path.basename(dwg_path))
        try:
            doc = acad.Documents.Open(dwg_path)
        except Exception as e:
            print("  打开失败，跳过。错误：", e)
            continue

        # 等图纸加载
        time.sleep(1.0)

        print(f"  将图层 {TARGET_LAYER_NAME} 的对象置于最前...")
        try:
            doc.SendCommand(lisp_cmd)
        except Exception as e:
            print("  发送命令失败，错误：", e)
            try:
                doc.Close(False)
            except Exception:
                pass
            continue

        # 等 DRAWORDER 命令完成，复杂图可适当加大
        time.sleep(3.0)

        try:
            doc.Save()
            doc.Close()
            print("  已处理并保存。")
        except Exception as e:
            print("  保存/关闭出错，错误：", e)
            try:
                doc.Close(False)
            except Exception:
                pass

    print("\n全部 DWG 处理完成，可随便打开几张检查一下效果。")


if __name__ == "__main__":
    main()
