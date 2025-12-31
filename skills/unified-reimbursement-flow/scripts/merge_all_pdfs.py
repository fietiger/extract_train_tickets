import os
import pypdfium2 as pdfium
import sys

def merge_pdfs(output_path):
    dest = pdfium.PdfDocument.new()
    
    # Order: Reimbursement -> Expense List -> Train -> Didi
    files = []
    if os.path.exists('费用报销单.pdf'): files.append('费用报销单.pdf')
    if os.path.exists('费用清单.pdf'): files.append('费用清单.pdf')
    
    for d in ['火车票', '滴滴出行电子发票及行程报销单']:
        if os.path.exists(d):
            for root, _, fs in os.walk(d):
                for f in fs:
                    if f.lower().endswith('.pdf'):
                        files.append(os.path.join(root, f))
    
    for f in files:
        src = pdfium.PdfDocument(f)
        dest.import_pages(src)
        src.close()
    dest.save(output_path)
    dest.close()

if __name__ == "__main__":
    merge_pdfs(sys.argv[1] if len(sys.argv) > 1 else '最终合并报销文件.pdf')
