import os
import re
import shutil
import subprocess
import sys

def which_abs(cmd, fallback):
    p = shutil.which(cmd)
    return p if p else fallback

LP        = which_abs("lp",        "/usr/bin/lp")
LPSTAT    = which_abs("lpstat",    "/usr/bin/lpstat")
LPOPTIONS = which_abs("lpoptions", "/usr/sbin/lpoptions")

def run(cmd):
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return (proc.stdout or "") + (proc.stderr or "")
    except Exception:
        return ""

DEVICE_RE = re.compile(r"^device for (.+?):", re.IGNORECASE)

def clean_printer_name(name: str) -> str:
    """去掉状态/冒号等尾巴，得到纯队列名"""
    name = name.strip()
    # device for NAME: ... → NAME
    m = DEVICE_RE.match(name)
    if m:
        name = m.group(1).strip()
    # 截断中文状态/英文状态关键词之前
    cut_markers = [
        " 正在", "正在", " 接受", "接受",  # 中文常见
        " accepting", " is ", " enabled", " disabled",  # 英文常见
    ]
    for mk in cut_markers:
        i = name.find(mk)
        if i > 0:
            name = name[:i].strip()
            break
    # 再截一次冒号（保险）
    if ":" in name:
        name = name.split(":", 1)[0].strip()
    return name

def list_printers():
    names = set()

    # 1) 首选：lpstat -v → “device for NAME:” 最稳定
    out = run([LPSTAT, "-v"])
    for line in out.splitlines():
        line = line.strip()
        m = DEVICE_RE.match(line)
        if m:
            names.add(m.group(1).strip())

    # 2) 兜底：lpstat -a / -p
    if not names:
        out = run([LPSTAT, "-a"])
        for line in out.splitlines():
            line = line.strip()
            # NAME 开头，后面是状态
            if line:
                names.add(clean_printer_name(line))

        out = run([LPSTAT, "-p"])
        for line in out.splitlines():
            line = line.strip()
            # printer NAME is ...
            if line.lower().startswith("printer "):
                parts = line.split()
                if len(parts) >= 2:
                    names.add(parts[1])

    # 3) 再从 lpoptions 收集（有些环境只在这里能拿到）
    out = run([LPOPTIONS])
    if out:
        toks = out.replace("default", "dest").split()
        for i, t in enumerate(toks):
            if t == "dest" and i + 1 < len(toks):
                names.add(toks[i + 1])

    # 默认打印机
    default_name = None
    out = run([LPSTAT, "-d"])
    if ":" in out:
        default_name = out.split(":", 1)[1].strip() or None
        default_name = clean_printer_name(default_name)

    # 清洗一遍名称集合
    clean = sorted({clean_printer_name(n) for n in names if n})
    # 把默认也放进去
    if default_name and default_name not in clean:
        clean.append(default_name)

    return clean, default_name

def batch_print_pdf(folder, printer, copies=1, two_sided=True):
    # 绝对路径存在性校验
    if not os.path.exists(LP):
        print(f"❌ 未找到 lp 命令：{LP}")
        sys.exit(1)

    # 清洗最终选择的打印机名
    printer = clean_printer_name(printer)

    opts = []
    if two_sided:
        opts += ["-o", "sides=two-sided-long-edge"]
    if copies and copies > 1:
        opts += ["-n", str(copies)]

    files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    files.sort()
    if not files:
        print("⚠️ 目录下没有 .pdf 文件：", folder)
        return

    for fname in files:
        pdf_path = os.path.join(folder, fname)
        if not os.path.isfile(pdf_path):
            print("⚠️ 跳过不存在文件：", pdf_path)
            continue
        print(f"🖨️ 打印 → {printer} ：{pdf_path}")
        cmd = [LP, "-d", printer] + opts + [pdf_path]
        ret = subprocess.run(cmd, capture_output=True, text=True)
        if ret.returncode != 0:
            # 直给原始输出，便于定位
            print("❌ lp 错误：", (ret.stdout or "") + (ret.stderr or ""))
        else:
            # 可选：显示返回的 job id
            msg = (ret.stdout or "").strip()
            if msg:
                print("✅", msg)

if __name__ == "__main__":
    folder = "/Users/asliujinhe/Downloads/export_docx_20250909_125826/"
    if not os.path.isdir(folder):
        print("❌ 目录不存在：", folder)
        sys.exit(1)

    printers, default_name = list_printers()
    if not printers and not default_name:
        print("❌ 未解析到打印机，请手动检查：")
        print("   ", LPSTAT, "-p -d")
        print("   ", LPSTAT, "-v")
        print("   ", LPOPTIONS)
        sys.exit(1)

    # 展示列表（默认打印机标注）
    unique = printers[:] or []
    if default_name and default_name not in unique:
        unique.append(default_name)

    print("可用打印机：")
    for i, p in enumerate(unique, 1):
        mark = " (默认)" if p == default_name else ""
        print(f"{i}. {p}{mark}")

    choice = input(f"请选择打印机编号（回车使用默认{(' '+default_name) if default_name else ''}）：").strip()
    if choice:
        try:
            printer = unique[int(choice) - 1]
        except Exception:
            print("❌ 选择无效")
            sys.exit(1)
    else:
        printer = default_name or unique[0]

    copies_in = input("份数（回车=1）：").strip()
    copies = int(copies_in) if copies_in.isdigit() and int(copies_in) > 0 else 1

    batch_print_pdf(folder, printer, copies=copies, two_sided=True)
    print("🎉 全部打印任务已提交（双面）")
