#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键串联公式转换流程：
1) 检查是否存在 MathType OLE 对象
2) 调用 01convert_equation_format_OLE_to_MathML.ps1
3) 调用 02convert_equation_format_MathML_to_OMML.ps1

说明：
- 为保证 Word 弹窗可人工处理，两个 PowerShell 子脚本均采用前台阻塞执行，不做超时强杀。
- 若第 1 步未发现 OLE 对象，默认停止；可通过 --continue-if-not-found 继续执行。
"""

from __future__ import annotations

import argparse
import os
import shlex
import shutil
import subprocess
import sys
from pathlib import Path


def _supports_color() -> bool:
    return sys.stdout.isatty() or bool(os.getenv("WT_SESSION")) or bool(os.getenv("TERM"))


USE_COLOR = _supports_color()
RESET = "\033[0m"
CYAN = "\033[96m"
BLUE = "\033[94m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
GRAY = "\033[90m"


def colorize(text: str, color: str) -> str:
    if not USE_COLOR:
        return text
    return f"{color}{text}{RESET}"


def print_step(text: str) -> None:
    print(colorize(text, CYAN))


def print_info(text: str) -> None:
    print(colorize(text, BLUE))


def print_notice(text: str) -> None:
    print(colorize(text, YELLOW))


def print_success(text: str) -> None:
    print(colorize(text, GREEN))


def print_error(text: str) -> None:
    print(colorize(text, RED), file=sys.stderr)


def print_command(text: str) -> None:
    print(colorize(text, GRAY))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="串联执行 OLE -> MathML -> OMML 转换流程"
    )
    parser.add_argument(
        "file",
        help="输入 Word 文件路径（.docx/.docm）",
    )
    parser.add_argument(
        "--out",
        help="最终输出路径（第 2 步输出）。不传则自动生成 *_MTConvert_OMML* 文件。",
    )
    parser.add_argument(
        "--continue-if-not-found",
        action="store_true",
        help="当第 1 步未发现 OLE 时继续执行后续步骤。",
    )
    parser.add_argument(
        "--show-check-details",
        action="store_true",
        help="显示 check_mathtype_ole.py 的匹配详情。",
    )
    parser.add_argument(
        "--macro1",
        default=None,
        help="第 1 步宏名称（不传则使用 01 脚本默认值）。",
    )
    parser.add_argument(
        "--macro2",
        default=None,
        help="第 2 步宏名称（不传则使用 02 脚本默认值）。",
    )
    parser.add_argument(
        "--show-word-step2",
        action="store_true",
        help="传递给第 2 步 ps1 的 -ShowWord 开关。",
    )
    return parser


def ensure_supported_doc(path: Path) -> None:
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"文件不存在：{path}")
    if path.suffix.lower() not in {".docx", ".docm"}:
        raise ValueError(f"仅支持 .docx/.docm，当前：{path}")


def unique_path(base: Path) -> Path:
    if not base.exists():
        return base
    i = 1
    while True:
        candidate = base.with_name(f"{base.stem}({i}){base.suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def choose_ps_exe() -> str:
    # 优先使用 pwsh，其次回退 Windows PowerShell
    pwsh = shutil.which("pwsh")
    if pwsh:
        return pwsh
    powershell = shutil.which("powershell")
    if powershell:
        return powershell
    raise RuntimeError("未找到 PowerShell 可执行文件（pwsh/powershell）")


def run_check(check_script: Path, doc_path: Path, show_details: bool) -> int:
    cmd = [sys.executable, str(check_script), str(doc_path)]
    if show_details:
        cmd.append("--show-details")

    print_step("[步骤 1/3] 检查 OLE")
    print_command("即将执行： " + " ".join(shlex.quote(c) for c in cmd))
    completed = subprocess.run(cmd)
    return completed.returncode


def run_ps_blocking(cmd: list[str], step_title: str) -> int:
    print_step(step_title)
    print_command("即将执行： " + " ".join(shlex.quote(c) for c in cmd))
    print_notice("提示：若 Word 出现弹窗，请手动点击确认，脚本会等待直到该步骤结束。")
    proc = subprocess.Popen(cmd)
    return proc.wait()


def main() -> int:
    args = build_parser().parse_args()
    project_dir = Path(__file__).resolve().parent

    check_script = project_dir / "check_mathtype_ole.py"
    step1_script = project_dir / "01convert_equation_format_OLE_to_MathML.ps1"
    step2_script = project_dir / "02convert_equation_format_MathML_to_OMML.ps1"

    for p in (check_script, step1_script, step2_script):
        if not p.exists():
            print_error(f"错误：缺少脚本 -> {p}")
            return 2

    try:
        source = Path(args.file).expanduser().resolve()
        ensure_supported_doc(source)
    except Exception as exc:
        print_error(f"错误：输入文件不合法，原因：{exc}")
        return 2

    check_rc = run_check(check_script, source, args.show_check_details)
    if check_rc == 2:
        print_error("错误：OLE 检查执行失败，流程中止。")
        return 2
    if check_rc == 0 and not args.continue_if_not_found:
        print_notice("未检测到 MathType OLE，对安全起见默认停止。")
        print_info("如需强制继续，请增加参数：--continue-if-not-found")
        return 0

    ps_exe = choose_ps_exe()

    # 第 1 步输出：与原脚本命名一致，默认 *_MTConvert*.docx/docm
    step1_out = unique_path(source.with_name(f"{source.stem}_MTConvert{source.suffix}"))
    step1_cmd = [
        ps_exe,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(step1_script),
        "-DocPath",
        str(source),
        "-OutPath",
        str(step1_out),
    ]
    if args.macro1:
        step1_cmd.extend(["-MacroName", args.macro1])

    step1_rc = run_ps_blocking(step1_cmd, "[步骤 2/3] OLE -> MathML")
    if step1_rc != 0:
        print_error(f"错误：步骤 2 执行失败，退出码={step1_rc}")
        return step1_rc
    if not step1_out.exists():
        print_error(f"错误：步骤 2 未生成输出文件：{step1_out}")
        return 2

    # 第 2 步输出：默认基于 step1_out 生成 *_OMML*
    if args.out:
        final_out = Path(args.out).expanduser().resolve()
    else:
        final_out = unique_path(
            step1_out.with_name(f"{step1_out.stem}_OMML{step1_out.suffix}")
        )
    final_out.parent.mkdir(parents=True, exist_ok=True)

    step2_cmd = [
        ps_exe,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(step2_script),
        "-DocxPath",
        str(step1_out),
        "-OutPath",
        str(final_out),
    ]
    if args.macro2:
        step2_cmd.extend(["-MacroName", args.macro2])
    if args.show_word_step2:
        step2_cmd.append("-ShowWord")

    step2_rc = run_ps_blocking(step2_cmd, "[步骤 3/3] MathML -> OMML")
    if step2_rc != 0:
        print_error(f"错误：步骤 3 执行失败，退出码={step2_rc}")
        return step2_rc
    if not final_out.exists():
        print_error(f"错误：步骤 3 未生成输出文件：{final_out}")
        return 2

    print_success(f"全部完成，最终输出：{final_out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
