import os
import csv
import threading
import configparser
from collections import defaultdict
from PyPDF2 import PdfReader
import docx2txt
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 추가 라이브러리 임포트
try:
    import pptx
except ImportError:
    pptx = None

try:
    import odf.opendocument
    import odf.text
    import odf.teletype
except ImportError:
    odf = None

# Windows 환경에서만 win32com.client를 사용할 수 있습니다.
if sys.platform == 'win32':
    import win32com.client
else:
    win32com = None

def get_pdf_page_count(file_path):
    try:
        with open(file_path, 'rb') as f:
            reader = PdfReader(f)
            return len(reader.pages)
    except Exception as e:
        print(f"Error processing PDF file {file_path}: {e}")
        return 0

def get_docx_page_count(file_path):
    try:
        text = docx2txt.process(file_path)
        # 대략적으로 1000자를 한 페이지로 가정하여 페이지 수 추정
        return len(text) // 1000 if text else 0
    except Exception as e:
        print(f"Error processing DOCX file {file_path}: {e}")
        return 0

def get_doc_page_count(file_path):
    if win32com is None:
        print(f"Cannot process .doc files on non-Windows platform: {file_path}")
        return 0
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path, ReadOnly=True)
        page_count = doc.ComputeStatistics(2)  # 2는 wdStatisticPages를 의미
        doc.Close()
        word.Quit()
        return page_count
    except Exception as e:
        print(f"Error processing DOC file {file_path}: {e}")
        return 0

def get_txt_page_count(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
            return len(text) // 1000 if text else 0
    except Exception as e:
        print(f"Error processing TXT file {file_path}: {e}")
        return 0

def get_pptx_page_count(file_path):
    if pptx is None:
        print(f"python-pptx library not installed, cannot process {file_path}")
        return 0
    try:
        prs = pptx.Presentation(file_path)
        return len(prs.slides)
    except Exception as e:
        print(f"Error processing PPTX file {file_path}: {e}")
        return 0

def get_odt_page_count(file_path):
    if odf is None:
        print(f"odfpy library not installed, cannot process {file_path}")
        return 0
    try:
        doc = odf.opendocument.load(file_path)
        allparas = odf.text.extractText(doc)
        return len(allparas) // 1000 if allparas else 0
    except Exception as e:
        print(f"Error processing ODT file {file_path}: {e}")
        return 0

def process_folder(folder_path, file_types, progress_callback=None, current_file_callback=None, add_file_result_callback=None):
    results = []
    folder_page_counts = defaultdict(int)
    total_pages = 0

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if not any(file.lower().endswith(ext) for ext in file_types):
                continue  # 선택된 파일 타입이 아니면 무시

            file_path = os.path.join(root, file)

            # 현재 처리 중인 파일명을 업데이트
            if current_file_callback:
                current_file_callback(file_path)

            page_count = 0
            file_type = ''
            file_name = file
            folder = root  # 폴더 경로

            if file.lower().endswith('.pdf'):
                page_count = get_pdf_page_count(file_path)
                file_type = 'pdf'
            elif file.lower().endswith('.docx'):
                page_count = get_docx_page_count(file_path)
                file_type = 'docx'
            elif file.lower().endswith('.doc'):
                page_count = get_doc_page_count(file_path)
                file_type = 'doc'
            elif file.lower().endswith('.txt'):
                page_count = get_txt_page_count(file_path)
                file_type = 'txt'
            elif file.lower().endswith('.pptx'):
                page_count = get_pptx_page_count(file_path)
                file_type = 'pptx'
            elif file.lower().endswith('.odt'):
                page_count = get_odt_page_count(file_path)
                file_type = 'odt'
            else:
                # 지원되지 않는 파일 타입은 무시
                continue

            if page_count is not None:
                result = {
                    'Type': file_type,
                    'Path': file_path,
                    'Name': file_name,
                    'Pages': page_count,
                    'Folder': folder
                }
                results.append(result)
                # 폴더별 페이지 수 합계 계산
                folder_page_counts[folder] += page_count
                # 총 페이지 수 계산
                total_pages += page_count

                # 파일별 결과를 GUI에 추가
                if add_file_result_callback:
                    add_file_result_callback(result)

            if progress_callback:
                progress_callback()

    # 작업이 완료되면 현재 파일명 표시를 지웁니다.
    if current_file_callback:
        current_file_callback('')  # 빈 문자열로 설정하여 표시 제거

    return results, folder_page_counts, total_pages

def write_results_to_csv(results, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Type', 'Path', 'Name', 'Pages']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        # 파일별 정보 작성
        for item in results:
            writer.writerow({
                'Type': item['Type'],
                'Path': item['Path'],
                'Name': item['Name'],
                'Pages': item['Pages']
            })

def write_summary_to_csv(folder_page_counts, total_pages, summary_file):
    with open(summary_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Folder', 'Total Pages'])
        for folder, page_sum in folder_page_counts.items():
            writer.writerow([folder, page_sum])
        writer.writerow(['Total', total_pages])

# 설정 파일 관리
CONFIG_FILE = 'config.cfg'

def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding='utf-8')
    return config

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

# GUI 구현
class PageCountApp:
    def __init__(self, root):
        self.root = root
        self.root.title("문서 페이지 수 계산기")

        # 창 크기 조정 가능하도록 설정
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        self.config = load_config()
        self.create_widgets()
        self.load_previous_settings()

    def create_widgets(self):
        # 메인 프레임 생성
        main_frame = tk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

        # 프레임 내 그리드 행열 구성 조정
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)

        # 루트 경로 선택
        self.path_label = tk.Label(main_frame, text="루트 경로:")
        self.path_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')

        self.path_entry = tk.Entry(main_frame)
        self.path_entry.grid(row=0, column=1, padx=5, pady=5, sticky='we')

        self.browse_button = tk.Button(main_frame, text="찾아보기", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        # 파일 타입 선택
        self.type_label = tk.Label(main_frame, text="파일 타입:")
        self.type_label.grid(row=1, column=0, padx=5, pady=5, sticky='ne')

        self.file_types = {
            '.doc': tk.BooleanVar(value=True),
            '.docx': tk.BooleanVar(value=True),
            '.pdf': tk.BooleanVar(value=True),
            '.txt': tk.BooleanVar(value=False),
            '.pptx': tk.BooleanVar(value=False),
            '.odt': tk.BooleanVar(value=False),
            # 필요한 경우 더 많은 파일 타입 추가
        }

        self.check_buttons = []
        row = 1
        col = 1
        for ext, var in self.file_types.items():
            chk = tk.Checkbutton(main_frame, text=ext, variable=var)
            chk.grid(row=row, column=col, sticky='w')
            self.check_buttons.append(chk)
            col += 1
            if col > 2:
                col = 1
                row += 1

        # 실행 버튼
        self.execute_button = tk.Button(main_frame, text="실행", command=self.start_processing)
        self.execute_button.grid(row=row+1, column=0, columnspan=3, pady=10)
        self.execute_button.config(width=20, height=2, font=('Arial', 12))

        # 진행 상황 표시
        self.progress = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
        self.progress.grid(row=row+2, column=0, columnspan=3, padx=5, pady=5, sticky='we')

        # 현재 처리 중인 파일 표시
        self.current_file_label = tk.Label(main_frame, text="현재 처리 중인 파일:")
        self.current_file_label.grid(row=row+3, column=0, padx=5, pady=5, sticky='ne')

        self.current_file_value = tk.Label(main_frame, text="", wraplength=400, justify='left')
        self.current_file_value.grid(row=row+3, column=1, columnspan=2, padx=5, pady=5, sticky='w')

        # 결과 표시 프레임 생성
        self.result_frame = tk.Frame(main_frame)
        self.result_frame.grid(row=row+4, column=0, columnspan=3, padx=5, pady=5, sticky='nsew')

        # 결과 프레임 내부 레이아웃 조정
        self.result_frame.columnconfigure(0, weight=1)
        self.result_frame.rowconfigure(1, weight=1)

        # 파일 목록 표시
        self.file_list_label = tk.Label(self.result_frame, text="파일별 페이지 수:")
        self.file_list_label.grid(row=0, column=0, sticky='w')

        self.file_list_text = tk.Text(self.result_frame, height=10)
        self.file_list_text.grid(row=1, column=0, sticky='nsew')

        # 스크롤바 추가
        self.file_list_scrollbar = tk.Scrollbar(self.result_frame, orient='vertical', command=self.file_list_text.yview)
        self.file_list_text.configure(yscrollcommand=self.file_list_scrollbar.set)
        self.file_list_scrollbar.grid(row=1, column=1, sticky='ns')

        # 구분선
        self.separator = ttk.Separator(self.result_frame, orient='horizontal')
        self.separator.grid(row=2, column=0, columnspan=2, sticky='we', pady=5)

        # 요약 결과 표시
        self.summary_label = tk.Label(self.result_frame, text="요약 결과:")
        self.summary_label.grid(row=3, column=0, sticky='w')

        self.summary_text = tk.Text(self.result_frame, height=10)
        self.summary_text.grid(row=4, column=0, sticky='nsew')

        # 스크롤바 추가
        self.summary_scrollbar = tk.Scrollbar(self.result_frame, orient='vertical', command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=self.summary_scrollbar.set)
        self.summary_scrollbar.grid(row=4, column=1, sticky='ns')

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)

    def load_previous_settings(self):
        if 'Settings' in self.config:
            settings = self.config['Settings']
            folder_path = settings.get('folder_path', '')
            self.path_entry.insert(0, folder_path)

            selected_types = settings.get('file_types', '')
            selected_types_list = selected_types.split(',')
            for ext, var in self.file_types.items():
                var.set(ext in selected_types_list)

    def save_current_settings(self):
        if 'Settings' not in self.config:
            self.config['Settings'] = {}
        settings = self.config['Settings']
        settings['folder_path'] = self.path_entry.get()

        selected_types = [ext for ext, var in self.file_types.items() if var.get()]
        settings['file_types'] = ','.join(selected_types)

        save_config(self.config)

    def start_processing(self):
        folder_path = self.path_entry.get()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("에러", "유효한 폴더 경로를 입력하세요.")
            return

        file_types = [ext for ext, var in self.file_types.items() if var.get()]

        if not file_types:
            messagebox.showerror("에러", "적어도 하나의 파일 타입을 선택하세요.")
            return

        # 설정 저장
        self.save_current_settings()

        # 진행 상황 초기화
        self.progress['value'] = 0
        self.file_list_text.delete(1.0, tk.END)
        self.summary_text.delete(1.0, tk.END)
        self.current_file_value.config(text="")

        # 별도의 스레드에서 작업 수행
        threading.Thread(target=self.process_files, args=(folder_path, file_types)).start()

    def process_files(self, folder_path, file_types):
        # 총 파일 개수 계산
        total_files = 0
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if any(file.lower().endswith(ext) for ext in file_types):
                    total_files += 1

        if total_files == 0:
            messagebox.showinfo("정보", "선택된 파일 타입의 파일이 없습니다.")
            return

        # 진행 상황 업데이트를 위한 변수 설정
        self.root.after(0, lambda: self.progress.config(maximum=total_files))
        processed_files = 0

        # 현재 처리 중인 파일명을 업데이트하기 위한 함수
        def update_current_file(file_path):
            def _update():
                self.current_file_value.config(text=file_path)
            self.root.after(0, _update)

        # 진행 상황 업데이트 함수
        def update_progress():
            nonlocal processed_files
            processed_files += 1
            def set_progress():
                self.progress['value'] = processed_files
            self.root.after(0, set_progress)

        # 파일별 결과를 추가하기 위한 함수
        def add_file_result(result):
            def _update():
                line = f"{result['Type']}, {result['Path']}, {result['Name']}, {result['Pages']} 페이지\n"
                self.file_list_text.insert(tk.END, line)
                self.file_list_text.see(tk.END)  # 자동 스크롤
            self.root.after(0, _update)

        # 파일 처리
        results, folder_page_counts, total_pages = process_folder(
            folder_path,
            file_types,
            progress_callback=update_progress,
            current_file_callback=update_current_file,
            add_file_result_callback=add_file_result
        )

        # 작업 완료 후 현재 파일명 표시 제거
        self.root.after(0, lambda: self.current_file_value.config(text=""))

        # 결과 CSV로 저장
        write_results_to_csv(results, 'output.csv')
        write_summary_to_csv(folder_page_counts, total_pages, 'summary.csv')

        # 요약 결과 표시
        def update_summary():
            self.summary_text.insert(tk.END, "폴더별 총 페이지 수:\n")
            for folder, page_sum in folder_page_counts.items():
                line = f"{folder}: {page_sum} 페이지\n"
                self.summary_text.insert(tk.END, line)

            self.summary_text.insert(tk.END, f"\n전체 총 페이지 수: {total_pages} 페이지\n")
        self.root.after(0, update_summary)

        messagebox.showinfo("완료", "작업이 완료되었습니다. 결과는 output.csv 및 summary.csv에 저장되었습니다.")

# 프로그램 실행
if __name__ == '__main__':
    root = tk.Tk()
    app = PageCountApp(root)
    root.mainloop()
    print("success!")
