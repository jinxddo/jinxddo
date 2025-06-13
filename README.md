import argparse
import os
import sys
import logging
import time
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import win32com.client
from PyPDF2 import PdfFileMerger
from glob import glob
from tqdm import tqdm

# --- 로깅 설정 ---
logger = None
def setup_logger():
    global logger
    try:
        log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        log_file_handler = logging.FileHandler('conversion.log', mode='w', encoding='utf-8')
        log_file_handler.setFormatter(log_formatter)
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)
        if logger.hasHandlers(): logger.handlers.clear()
        logger.addHandler(log_file_handler)
    except Exception as e:
        print(f"!!! [경고] 로그 파일을 생성할 수 없습니다. 이유: {e}")
        logger = None

def log_and_print(message, level=logging.INFO):
    print(message, flush=True)
    if logger:
        if level == logging.INFO: logger.info(message)
        elif level == logging.ERROR: logger.error(message)

# --- 기능 함수 정의 ---
def create_cover(target_docx_path, base_doc_path, f_name, yy, mm):
    """ 표지 생성: 다양한 문서 규격에 대응하는 최종 'Plan A, Plan B' 로직 적용 """
    log_and_print(f"\n[{f_name}] 1. 표지 생성 작업 시작")
    try:
        log_and_print(f"  -> 표지 생성을 위해 원본 문서를 엽니다: {os.path.basename(target_docx_path)}")
        docx = Document(target_docx_path)
        log_and_print("  -> 문서에서 제목/버전 정보를 추출합니다.")
        
        doc_context = []
        
        # ================== 핵심 수정: Plan A / Plan B 제목 추출 로직 적용 ==================
        # Plan A: 사용자님의 원래 로직으로 빠르게 먼저 시도
        is_title_in_para = False
        # 첫 줄에 빈 줄이 있을 수 있으므로, 상위 3개 문단을 확인
        for p in docx.paragraphs[:3]:
            if p.text.strip().startswith("3GPP"):
                is_title_in_para = True
                break
        
        if is_title_in_para:
            log_and_print("  -> [Plan A] '일반 문단' 형식으로 제목 추출을 시도합니다.")
            for p in docx.paragraphs:
                if p.text.strip().startswith("The present"): break
                if p.text.strip(): doc_context.append(p.text.strip())
        
        # Plan A의 문단 방식이 실패했고, 표가 존재할 경우
        if not doc_context and docx.tables:
            log_and_print("  -> [Plan A] '표' 형식으로 제목 추출을 시도합니다.")
            try:
                # 첫 번째 표 형식 시도 (하나의 큰 표)
                title_tables = docx.tables
                temp_context = []
                temp_context.append(title_tables[0].cell(0,0).text)
                temp_context.append(title_tables[0].cell(1,0).text)
                for para in title_tables[0].cell(2,0).paragraphs: temp_context.append(para.text)
                temp_context.extend([title_tables[0].cell(3,0).text, title_tables[0].cell(4,0).text])
                # 유효성 검사 후 할당
                if temp_context and temp_context[0].strip():
                    doc_context = [line.strip() for line in temp_context if line.strip()]
            except Exception:
                doc_context = [] # 실패 시 다음 계획을 위해 초기화
        
        # Plan B: Plan A가 실패하여 doc_context가 여전히 비어있을 경우, 키워드 탐색으로 다시 시도
        if not doc_context:
            log_and_print("  -> [Plan A 실패] 키워드 기반 전체 탐색(Plan B)으로 전환합니다.")
            all_text_lines = []
            for para in docx.paragraphs:
                if para.text.strip(): all_text_lines.append(para.text.strip())
            for table in docx.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip(): all_text_lines.append(cell.text.strip())
            
            doc_number_line = next((line for line in all_text_lines if "3GPP TS" in line or "3GPP TR" in line), None)
            start_index = next((i for i, line in enumerate(all_text_lines) if "3rd Generation" in line), -1)
            end_index = next((i for i, line in enumerate(all_text_lines) if start_index != -1 and i >= start_index and "Release" in line and line.strip().endswith(')')), -1)
            
            if doc_number_line and start_index != -1 and end_index != -1:
                log_and_print(f"  -> [Plan B 성공] 상세 제목 블록을 정확히 추출했습니다.")
                doc_context.append(doc_number_line)
                doc_context.extend(all_text_lines[start_index : end_index + 1])
        # =================================================================================

        # --- 최종 확인 ---
        if not doc_context or not doc_context[0].strip():
            log_and_print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", level=logging.ERROR)
            log_and_print(f"!!! [최종 실패] '{f_name}' 문서에서 제목 정보를 추출하지 못했습니다.", level=logging.ERROR)
            log_and_print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", level=logging.ERROR)
            return False, ""
        
        # --- 이하 모든 코드는 보내주신 원본과 동일합니다 ---
        docx_cover = Document(base_doc_path)
        tables = docx_cover.tables
        
        tables[0].cell(1, 1).paragraphs[0].text = 'TTAT.3G-' + doc_context[0][8:len(doc_context[0]) - 10]
        tables[0].cell(1, 1).paragraphs[0].style.font.name = '돋움'
        tables[0].cell(1, 1).paragraphs[0].style.font.size = Pt(15)
        tables[0].cell(1, 1).paragraphs[0].style.font.bold = True
        tables[0].cell(1, 3).paragraphs[0].text = f'제정일: {yy}.{mm}.'
        tables[0].cell(1, 3).paragraphs[0].style.font.name = '돋움'
        tables[0].cell(1, 3).paragraphs[0].style.font.size = Pt(15)
        tables[0].cell(1, 3).paragraphs[0].style.font.bold = True
        
        final_lines_to_print = []
        for line_text in doc_context[1:]:
            processed_text = line_text.replace(';', ';\n')
            final_lines_to_print.extend(s.strip() for s in processed_text.split('\n') if s.strip())
        
        num_lines = len(final_lines_to_print)
	log_and_print(f"    -> [정보] 최종적으로 표지에 들어갈 상세 제목은 총 {num_lines}줄 입니다.")

        uniform_font_size = Pt(20)
        if num_lines >= 9: uniform_font_size = Pt(15)
        elif num_lines >= 8: uniform_font_size = Pt(17)

        target_cell = tables[0].cell(4, 3).tables[0].cell(0, 0)
        target_paragraphs = target_cell.paragraphs
        
        for idx, line in enumerate(final_lines_to_print):
            try:
                p = target_paragraphs[idx]
                p.text = line
            except IndexError:
                p = target_cell.add_paragraph(line)
            
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.style.font.name = '돋움'
            p.style.font.size = uniform_font_size
            p.style.font.bold = True
        
        output_cover_path = os.path.join(os.path.dirname(target_docx_path), f'!{f_name}.docx')
        docx_cover.save(output_cover_path)
        log_and_print(f"  -> [성공] 생성된 표지를 파일로 저장했습니다: {os.path.basename(output_cover_path)}")
        return True, os.path.basename(output_cover_path)
    except Exception as e:
        log_and_print(f"  -> [오류] 표지 생성 중 예기치 못한 문제 발생: {e}", level=logging.ERROR)
        return False, ""

def process_folder(current_folder_path, base_doc_path, output_root, f_name, yy, mm):
    """ 한 폴더의 모든 작업을 관리하고, Word 인스턴스를 한번만 실행하여 속도를 최적화합니다. """
    log_and_print(f"\n{'='*25} [{f_name}] 작업 시작 {'='*25}")
    os.system("taskkill /f /im winword.exe > nul 2> nul")
    word_instance = None
    try:
        log_and_print("  -> MS Word를 실행합니다. (이 폴더의 모든 작업에 재사용됩니다)")
        word_instance = win32com.client.Dispatch('Word.Application')
        word_instance.visible = False

        all_docs_initial = glob(os.path.join(current_folder_path, '*.doc*'))
        main_content_file = next((f for f in all_docs_initial if not os.path.basename(f).startswith(('!', '~'))), None)
        
        if not main_content_file:
            log_and_print(f"  -> [경고] 처리할 원본 문서가 없어 이 폴더를 건너뜁니다.", level=logging.WARNING)
            return True

        target_for_cover = main_content_file
        if main_content_file.lower().endswith('.doc'):
            doc_to_convert = word_instance.Documents.Open(main_content_file, OpenAndRepair=True)
            doc_to_convert.SaveAs(f"{main_content_file}x", 12)
            doc_to_convert.Close(SaveChanges=0)
            target_for_cover += 'x'
        
        cover_success, _ = create_cover(target_for_cover, base_doc_path, f_name, yy, mm)
        if not cover_success: return False

        pdf_name = f"TTAT.3G-{f_name}"
        try:
            v1,v2,v3='','','';pnp="TTAT.3G-"
            if len(f_name)<10:(vc1,vc2,vc3)=ord(f_name[6]),ord(f_name[7]),ord(f_name[8]);v1,v2,v3=(str(c-87)if c>90 else str(c-48)for c in[vc1,vc2,vc3]);pdf_name=f"{pnp}{f_name[0:2]}.{f_name[2:5]}V{v1}.{v2}.{v3}"
            else:(vc1,vc2,vc3)=ord(f_name[8]),ord(f_name[9]),ord(f_name[10]);v1,v2,v3=(str(c-87)if c>90 else str(c-48)for c in[vc1,vc2,vc3]);pdf_name=f"{pnp}{f_name[0:2]}.{f_name[2:7]}V{v1}.{v2}.{v3}"
        except IndexError: pass
        
        all_word_files_to_process = glob(os.path.join(current_folder_path, '*.doc*'))
        generated_pdfs, failed_files = [], []
        
        for filepath in all_word_files_to_process:
            filename = os.path.basename(filepath)
            if filename.startswith('~'): continue
            if main_content_file.lower().endswith('.doc') and filepath.lower() == main_content_file.lower(): continue
                
            if filename.lower().endswith(('.doc', '.docx')):
                try:
                    outputFile = os.path.splitext(filepath)[0] + ".pdf"
                    log_and_print(f"  -> [PDF 변환] 시도: {filename}")
                    doc = word_instance.Documents.Open(filepath, OpenAndRepair=True)
                    doc.ExportAsFixedFormat(OutputFileName=outputFile, ExportFormat=17)
                    
                    # --- 핵심 수정: 파일이 실제로 생성되었는지 확인하는 로직 ---
                    log_and_print(f"    -> 파일 생성 확인 중... (최대 60초 대기)")
                    timeout = 60
                    file_ready = False
                    while timeout > 0:
                        if os.path.exists(outputFile) and os.path.getsize(outputFile) > 0:
                            file_ready = True
                            break
                        time.sleep(1)
                        timeout -= 1
                    
                    if not file_ready:
                        raise Exception("PDF 변환 후 파일이 생성되지 않음 (타임아웃)")
                    
                    doc.Close(SaveChanges=0)
                    generated_pdfs.append(outputFile)
                    log_and_print(f"    -> [성공] 변환 완료: {os.path.basename(outputFile)}")
                except Exception as e:
                    log_and_print(f"    -> [!!! 변환 실패 !!!] 파일: {filename}, 이유: {e}", level=logging.ERROR)
                    failed_files.append(filename)
                    if 'doc' in locals() and doc: doc.Close(SaveChanges=0)
        
        if failed_files:
            log_and_print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!",level=logging.ERROR)
            log_and_print(f"!!! [오류 감지] '{f_name}' 폴더 작업 중 변환 실패 파일이 있어 병합을 건너뛰고 프로그램을 중단합니다.",level=logging.ERROR)
            for f in failed_files: log_and_print(f"!!!   - {f}",level=logging.ERROR)
            log_and_print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!",level=logging.ERROR)
            return False
        
        merger = PdfFileMerger()
        generated_pdfs.sort(key=lambda x: not os.path.basename(x).startswith('!'))
        for pdf in generated_pdfs: merger.append(pdf)
        merger.write(os.path.join(output_root, f"{pdf_name}.pdf"))
        merger.close()
        log_and_print(f"  -> [병합 성공] 최종 파일 저장: {pdf_name}.pdf")
        return True
    except Exception as e:
        log_and_print(f"  -> [치명적 오류] '{f_name}' 폴더 처리 중 예기치 못한 문제 발생: {e}", level=logging.ERROR)
        return False
    finally:
        if word_instance:
            word_instance.Quit()
            log_and_print("  -> MS Word 프로세스를 종료했습니다.")

def main():
    setup_logger()
    if os.name == 'nt' and sys.stdout.encoding.lower().find('utf') == -1:
        log_and_print("!!! [경고] CMD 창 인코딩 문제. 'chcp 65001'을 먼저 실행해주세요. !!!")
    
    parser = argparse.ArgumentParser(description="TTA-3GPP 문서 변환 자동화 스크립트")
    parser.add_argument('--path', '-P', type=str, default='.')
    parser.add_argument('--input', '-i', required=True)
    parser.add_argument('--output', '-o', required=True)
    parser.add_argument('--base_input', '-B', required=True)
    parser.add_argument('--year', '-Y', required=True)
    parser.add_argument('--month', '-M', required=True)
    args = parser.parse_args()

    try:
        input_root = os.path.join(args.path, args.input)
        output_root = os.path.join(args.path, args.output)
        base_doc_path = os.path.join(args.path, args.base_input)
        
        folder_list = [d for d in os.listdir(input_root) if os.path.isdir(os.path.join(input_root, d))]
        if not folder_list:
            log_and_print(f"정보: 처리할 하위 폴더가 없습니다 -> '{input_root}'")
            return

        log_and_print(f"총 {len(folder_list)}개의 폴더에 대한 작업을 시작합니다.")
        progress_bar = tqdm(folder_list, desc="전체 진행률", unit="폴더")
        
        for folder_name in progress_bar:
            success = process_folder(os.path.join(input_root, folder_name), base_doc_path, output_root, folder_name, args.year, args.month)
            if not success:
                log_and_print("\n!!! 작업이 중단되었습니다. 위에 표시된 오류를 확인해주세요. !!!", level=logging.ERROR)
                break 

        progress_bar.close()
        log_and_print("\n★★★★★ 작업이 종료되었습니다. ★★★★★")
    except Exception as e:
        log_and_print(f"\n\n[치명적 오류] 스크립트 실행 중 문제 발생: {e}", level=logging.ERROR)

if __name__ == '__main__':
    main()
