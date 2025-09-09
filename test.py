import os
import subprocess

def word_to_pdf_mac(word_file, pdf_file):
    script = f'''
    tell application "Microsoft Word"
        open POSIX file "{word_file}"
        set theDoc to active document
        save as theDoc file name (POSIX file "{pdf_file}") file format format PDF
        close theDoc saving no
    end tell
    '''
    subprocess.run(["osascript", "-e", script], check=True)

def batch_convert(folder):
    for name in os.listdir(folder):
        if name.lower().endswith((".docx", ".doc")):
            word_path = os.path.join(folder, name)
            pdf_path = os.path.join(folder, os.path.splitext(name)[0] + ".pdf")
            print(f"Converting: {word_path} -> {pdf_path}")
            word_to_pdf_mac(word_path, pdf_path)

if __name__ == "__main__":
    folder = "/Users/asliujinhe/Downloads/export_docx_20250909_125826/"
    batch_convert(folder)
    print("✅ 批量转换完成！")
