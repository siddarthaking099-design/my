import os
import tempfile
import shutil
from flask import Flask, request, send_file, jsonify, after_this_request
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from docx2pdf import convert
import fitz  # PyMuPDF
import pandas as pd
import tabula
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image
import img2pdf
import base64
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import uuid
import io

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'xlsx', 'png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename) or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        # Save uploaded PDF
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        # Convert PDF to Word
        output_filename = f"{filename[:-4]}.docx"
        docx_path = os.path.join(temp_dir, output_filename)
        
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        
        # Copy to a new temp file for sending
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        shutil.copy2(docx_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename) or not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Invalid file type. Please upload a DOCX file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        # Save uploaded Word document
        filename = secure_filename(file.filename)
        docx_path = os.path.join(temp_dir, filename)
        file.save(docx_path)
        
        # Convert Word to PDF
        output_filename = f"{filename[:-5]}.pdf"
        pdf_path = os.path.join(temp_dir, output_filename)
        
        convert(docx_path, pdf_path)
        
        # Copy to a new temp file for sending
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(pdf_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        output_filename = f"{filename[:-4]}.xlsx"
        excel_path = os.path.join(temp_dir, output_filename)
        
        # Extract tables using tabula
        try:
            tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        except:
            # Fallback to pdfplumber for text extraction
            with pdfplumber.open(pdf_path) as pdf:
                all_text = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text.append(text.split('\n'))
                
                if all_text:
                    # Create DataFrame from text
                    max_cols = max(len(row) for page in all_text for row in page if row.strip())
                    data = []
                    for page in all_text:
                        for row in page:
                            if row.strip():
                                cols = row.split()
                                cols.extend([''] * (max_cols - len(cols)))
                                data.append(cols[:max_cols])
                    tables = [pd.DataFrame(data)] if data else [pd.DataFrame([['No data extracted']])]
                else:
                    tables = [pd.DataFrame([['No data extracted']])]
        
        # Create Excel workbook
        wb = Workbook()
        wb.remove(wb.active)
        
        for i, table in enumerate(tables):
            ws = wb.create_sheet(title=f'Table_{i+1}')
            for r in dataframe_to_rows(table, index=False, header=True):
                ws.append(r)
        
        wb.save(excel_path)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        shutil.copy2(excel_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'Invalid file type. Please upload an XLSX file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        excel_path = os.path.join(temp_dir, filename)
        file.save(excel_path)
        
        output_filename = f"{filename[:-5]}.pdf"
        pdf_path = os.path.join(temp_dir, output_filename)
        
        # Read Excel file
        excel_data = pd.read_excel(excel_path, sheet_name=None)
        
        # Create PDF using reportlab
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)
        elements = []
        
        for sheet_name, df in excel_data.items():
            # Convert DataFrame to list for table
            data = [df.columns.tolist()] + df.fillna('').astype(str).values.tolist()
            
            # Create table
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            elements.append(table)
        
        doc.build(elements)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(pdf_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/scan-pdf', methods=['POST'])
def scan_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        doc = fitz.open(pdf_path)
        pages = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            mat = fitz.Matrix(1.0, 1.0)  # Preview quality
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            
            pages.append({
                'page_num': page_num + 1,
                'preview': f'data:image/png;base64,{img_base64}'
            })
        
        doc.close()
        
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        return jsonify({'pages': pages})
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Scan failed: {str(e)}'}), 500

@app.route('/api/convert-pages', methods=['POST'])
def convert_pages():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    pages = request.form.get('pages', '').split(',')
    output_format = request.form.get('format', 'png')
    
    if not pages or pages == ['']:
        return jsonify({'error': 'No pages selected'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        doc = fitz.open(pdf_path)
        images = []
        
        for page_str in pages:
            page_num = int(page_str) - 1
            if 0 <= page_num < len(doc):
                page = doc.load_page(page_num)
                mat = fitz.Matrix(2.0, 2.0)  # High quality
                pix = page.get_pixmap(matrix=mat)
                
                img_data = pix.tobytes(output_format)
                img_base64 = base64.b64encode(img_data).decode('utf-8')
                
                images.append({
                    'filename': f'page_{page_num+1:03d}.{output_format}',
                    'data': img_base64
                })
        
        doc.close()
        
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        return jsonify({'images': images})
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/image-to-pdf', methods=['POST'])
def image_to_pdf():
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No files selected'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        # Process images
        image_paths = []
        for file in files:
            if file.filename and any(file.filename.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp']):
                filename = secure_filename(file.filename)
                img_path = os.path.join(temp_dir, filename)
                file.save(img_path)
                
                # Convert to RGB if needed
                with Image.open(img_path) as img:
                    if img.mode != 'RGB':
                        rgb_img = img.convert('RGB')
                        rgb_path = os.path.join(temp_dir, f'rgb_{filename}')
                        rgb_img.save(rgb_path, 'JPEG', quality=95)
                        image_paths.append(rgb_path)
                    else:
                        image_paths.append(img_path)
        
        if not image_paths:
            return jsonify({'error': 'No valid image files found'}), 400
        
        # Create PDF
        pdf_filename = 'converted_images.pdf'
        pdf_path = os.path.join(temp_dir, pdf_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(img2pdf.convert(image_paths))
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(pdf_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/merge-pdf', methods=['POST'])
def merge_pdf():
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No files selected'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        merger = PdfMerger()
        
        for file in files:
            if file.filename and file.filename.lower().endswith('.pdf'):
                filename = secure_filename(file.filename)
                pdf_path = os.path.join(temp_dir, filename)
                file.save(pdf_path)
                merger.append(pdf_path)
        
        merged_filename = 'merged_document.pdf'
        merged_path = os.path.join(temp_dir, merged_filename)
        
        merger.write(merged_path)
        merger.close()
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(merged_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=merged_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Merge failed: {str(e)}'}), 500

@app.route('/api/split-pdf', methods=['POST'])
def split_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        with open(pdf_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            pages = []
            
            for i in range(len(pdf_reader.pages)):
                writer = PdfWriter()
                writer.add_page(pdf_reader.pages[i])
                
                page_filename = f'page_{i+1:03d}.pdf'
                page_path = os.path.join(temp_dir, page_filename)
                
                with open(page_path, 'wb') as out:
                    writer.write(out)
                
                with open(page_path, 'rb') as page_file:
                    page_data = page_file.read()
                    page_base64 = base64.b64encode(page_data).decode('utf-8')
                
                pages.append({
                    'filename': page_filename,
                    'data': page_base64
                })
        
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        return jsonify({'pages': pages})
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Split failed: {str(e)}'}), 500

@app.route('/api/rotate-pdf', methods=['POST'])
def rotate_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    rotation = int(request.form.get('rotation', 90))
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        with open(pdf_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            writer = PdfWriter()
            
            for page in pdf_reader.pages:
                page.rotate(rotation)
                writer.add_page(page)
        
        rotated_filename = f'rotated_{filename}'
        rotated_path = os.path.join(temp_dir, rotated_filename)
        
        with open(rotated_path, 'wb') as out:
            writer.write(out)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(rotated_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=rotated_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Rotation failed: {str(e)}'}), 500

@app.route('/api/delete-pages', methods=['POST'])
def delete_pages():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    pages_to_delete = request.form.get('pages', '')
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        # Parse pages to delete (e.g., "1,3,5-7")
        pages_to_remove = set()
        for part in pages_to_delete.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages_to_remove.update(range(start-1, end))
            else:
                pages_to_remove.add(int(part)-1)
        
        with open(pdf_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            writer = PdfWriter()
            
            for i, page in enumerate(pdf_reader.pages):
                if i not in pages_to_remove:
                    writer.add_page(page)
        
        modified_filename = f'modified_{filename}'
        modified_path = os.path.join(temp_dir, modified_filename)
        
        with open(modified_path, 'wb') as out:
            writer.write(out)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(modified_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=modified_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Page deletion failed: {str(e)}'}), 500

@app.route('/api/compress-pdf', methods=['POST'])
def compress_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    compression_level = request.form.get('compression_level', 'medium')
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        # Aggressive compression settings based on level
        if compression_level == 'high':
            image_quality = 30
            max_width = 1200
        elif compression_level == 'medium':
            image_quality = 50
            max_width = 1600
        else:  # low
            image_quality = 70
            max_width = 2000
        
        # Multiple pass compression
        doc = fitz.open(pdf_path)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # Aggressive image optimization
                    img_temp = io.BytesIO(image_bytes)
                    pil_image = Image.open(img_temp)
                    
                    # Resize large images
                    if pil_image.width > max_width or pil_image.height > max_width:
                        ratio = min(max_width / pil_image.width, max_width / pil_image.height)
                        new_size = (int(pil_image.width * ratio), int(pil_image.height * ratio))
                        pil_image = pil_image.resize(new_size, Image.LANCZOS)
                    
                    # Convert to RGB
                    if pil_image.mode in ('RGBA', 'LA'):
                        background = Image.new('RGB', pil_image.size, (255, 255, 255))
                        if pil_image.mode == 'RGBA':
                            background.paste(pil_image, mask=pil_image.split()[3])
                        else:
                            background.paste(pil_image, mask=pil_image.split()[1])
                        pil_image = background
                    elif pil_image.mode != 'RGB':
                        pil_image = pil_image.convert('RGB')
                    
                    # Aggressive JPEG compression
                    optimized_img = io.BytesIO()
                    pil_image.save(optimized_img, format='JPEG', quality=image_quality, optimize=True, progressive=True)
                    optimized_img_bytes = optimized_img.getvalue()
                    
                    # Replace image in PDF
                    doc._replaceImage(xref, stream=optimized_img_bytes)
                    
                except Exception:
                    continue
        
        compressed_filename = f'compressed_{filename}'
        compressed_path = os.path.join(temp_dir, compressed_filename)
        
        # Maximum compression settings
        doc.save(compressed_path, garbage=4, deflate=True, clean=True)
        doc.close()
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(compressed_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=compressed_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Compression failed: {str(e)}'}), 500

@app.route('/api/protect-pdf', methods=['POST'])
def protect_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    user_password = request.form.get('user_password', '')
    owner_password = request.form.get('owner_password', '')
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    if not user_password:
        return jsonify({'error': 'Password is required'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        # Read and protect PDF
        with open(pdf_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            pdf_writer = PdfWriter()
            
            # Add all pages
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
            
            # Encrypt with password
            pdf_writer.encrypt(
                user_password=user_password,
                owner_password=owner_password or user_password
            )
        
        protected_filename = f'protected_{filename}'
        protected_path = os.path.join(temp_dir, protected_filename)
        
        with open(protected_path, 'wb') as out:
            pdf_writer.write(out)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(protected_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=protected_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Protection failed: {str(e)}'}), 500

@app.route('/api/unlock-pdf', methods=['POST'])
def unlock_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    password = request.form.get('password', '')
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        
        # Try to unlock PDF
        with open(pdf_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            
            if not pdf_reader.is_encrypted:
                # PDF is not encrypted, just return it as unlocked
                unlocked_filename = f'unlocked_{filename}'
                unlocked_path = os.path.join(temp_dir, unlocked_filename)
                shutil.copy2(pdf_path, unlocked_path)
                
                final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                shutil.copy2(unlocked_path, final_temp.name)
                final_temp.close()
                
                @after_this_request
                def cleanup(response):
                    try:
                        if temp_dir and os.path.exists(temp_dir):
                            shutil.rmtree(temp_dir)
                        if os.path.exists(final_temp.name):
                            os.unlink(final_temp.name)
                    except:
                        pass
                    return response
                
                return send_file(
                    final_temp.name,
                    as_attachment=True,
                    download_name=unlocked_filename,
                    mimetype='application/pdf'
                )
            
            # Try with provided password or attempt common passwords
            passwords_to_try = []
            if password:
                passwords_to_try.append(password)
            else:
                # Common passwords to try
                passwords_to_try.extend([
                    '', '123456', 'password', '123456789', '12345678',
                    'abc123', 'Password', '123123', 'admin', 'user'
                ])
            
            unlocked = False
            used_password = None
            
            for pwd in passwords_to_try:
                if pdf_reader.decrypt(pwd):
                    unlocked = True
                    used_password = pwd
                    break
            
            if not unlocked:
                return jsonify({'error': 'Could not unlock PDF. Invalid password or encryption too strong.'}), 400
            
            # Create unlocked PDF
            pdf_writer = PdfWriter()
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
        
        unlocked_filename = f'unlocked_{filename}'
        unlocked_path = os.path.join(temp_dir, unlocked_filename)
        
        with open(unlocked_path, 'wb') as out:
            pdf_writer.write(out)
        
        final_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        shutil.copy2(unlocked_path, final_temp.name)
        final_temp.close()
        
        @after_this_request
        def cleanup(response):
            try:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if os.path.exists(final_temp.name):
                    os.unlink(final_temp.name)
            except:
                pass
            return response
        
        return send_file(
            final_temp.name,
            as_attachment=True,
            download_name=unlocked_filename,
            mimetype='application/pdf'
        )
    
    except Exception as e:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return jsonify({'error': f'Unlock failed: {str(e)}'}), 500

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'healthy', 'message': 'PDF conversion service is running'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)