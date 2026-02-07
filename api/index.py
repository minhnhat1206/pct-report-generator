
from flask import Flask, render_template, request, jsonify, send_from_directory, send_file
import os
import zipfile
import io
import sys
from werkzeug.utils import secure_filename
import logging
import mammoth
from docx import Document

# Add root and current dir to path for Vercel
current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.dirname(current_dir)
for d in [current_dir, root_dir]:
    if d not in sys.path:
        sys.path.append(d)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

try:
    from report_grade_10 import generate_grade_10_reports
    from report_grade_11 import generate_grade_11_reports
    logger.info("Successfully imported report modules")
except Exception as e:
    logger.error(f"FAILED to import report modules: {e}")
    # Define dummy functions so the app still starts and we can see logs
    def generate_grade_10_reports(*args, **kwargs): raise e
    def generate_grade_11_reports(*args, **kwargs): raise e

app = Flask(__name__)

# Use /tmp for writable storage on Vercel, otherwise use local dirs
if os.environ.get('VERCEL'):
    BASE_TEMP = '/tmp'
    app.template_folder = '../templates'
    app.static_folder = '../static'
else:
    BASE_TEMP = current_dir

app.config['UPLOAD_FOLDER'] = os.path.join(BASE_TEMP, 'uploads')
app.config['GRADE_10_DIR'] = os.path.join(BASE_TEMP, 'Grade_10')
app.config['GRADE_11_DIR'] = os.path.join(BASE_TEMP, 'Grade_11')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GRADE_10_DIR'], exist_ok=True)
os.makedirs(app.config['GRADE_11_DIR'], exist_ok=True)

logging.basicConfig(level=logging.INFO)

import math

def clean_nans(value):
    """Recursively replace NaNs and Infinity with None (which becomes null in JSON)"""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    elif isinstance(value, dict):
        return {k: clean_nans(v) for k, v in value.items()}
    elif isinstance(value, list):
        return [clean_nans(v) for v in value]
    return value

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/debug')
def debug():
    files = os.listdir(root_dir)
    api_files = os.listdir(current_dir)
    return jsonify({
        'root_dir': root_dir,
        'current_dir': current_dir,
        'sys_path': sys.path,
        'root_files': files,
        'api_files': api_files,
        'vercel_env': os.environ.get('VERCEL', 'False')
    })

@app.route('/generate', methods=['POST'])
def generate():
    try:
        # Get Grade 10 inputs
        try:
            week_10 = int(request.form['week_10'])
            ielts_10 = int(request.form['ielts_10'])
            vstep_10 = int(request.form['vstep_10'])
            file_10 = request.files['file_10']
            
            # New config inputs for Grade 10
            course_vstep_10 = request.form.get('course_vstep_10', "Practical English A2-B2")
            course_ielts_10 = request.form.get('course_ielts_10', "Practical English A1")
            classes_10_str = request.form.get('classes_10', "10E1, 10E2, 10E3, 10E4")
            target_classes_10 = [cls.strip() for cls in classes_10_str.split(',') if cls.strip()]

        except (ValueError, KeyError) as e:
            return jsonify({'success': False, 'message': f'Lỗi dữ liệu đầu vào Khối 10: {str(e)}'})

        # Get Grade 11 inputs
        try:
            week_11 = int(request.form['week_11'])
            ielts_11 = int(request.form['ielts_11'])
            vstep_11 = int(request.form['vstep_11'])
            file_11 = request.files['file_11']

            # New config inputs for Grade 11
            course_vstep_11 = request.form.get('course_vstep_11', "Practical English A2-B2")
            course_ielts_11 = request.form.get('course_ielts_11', "Practical English B2 & IELTS A2-B1")
            classes_11_str = request.form.get('classes_11', "11E1, 11E2, 11E3, 11E4, 11V1, 11V2, 11V3, 11V4, 11V5, 11V6")
            target_classes_11 = [cls.strip() for cls in classes_11_str.split(',') if cls.strip()]

        except (ValueError, KeyError) as e:
            return jsonify({'success': False, 'message': f'Lỗi dữ liệu đầu vào Khối 11: {str(e)}'})

        # Check if files are provided
        timesheet_file = request.files.get('timesheet_file')
        if not file_10 or not file_11 or not timesheet_file:
            return jsonify({'success': False, 'message': 'Vui lòng cung cấp cả 3 file dữ liệu (Grade 10, 11 và Timesheet).'})

        # Save uploaded files
        path_10 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file_10.filename))
        path_11 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file_11.filename))
        path_timesheet = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(timesheet_file.filename))
        
        file_10.save(path_10)
        file_11.save(path_11)
        timesheet_file.save(path_timesheet)

        # Get lesson totals
        total_ielts_10 = int(request.form.get('total_ielts_10', 32))
        total_vstep_10 = int(request.form.get('total_vstep_10', 57))
        total_ielts_11 = int(request.form.get('total_ielts_11', 54))
        total_vstep_11 = int(request.form.get('total_vstep_11', 52))

        # Generate Reports
        generated_10 = []
        generated_11 = []
        stats_10 = None
        stats_11 = None

        try:
            out_dir_10 = os.path.join(app.config['GRADE_10_DIR'], f'Grade_10_Week {week_10}')
            generated_10, stats_10 = generate_grade_10_reports(
                week_10, vstep_10, ielts_10, path_10, output_dir=out_dir_10,
                target_classes_list=target_classes_10,
                course_vstep=course_vstep_10,
                course_ielts=course_ielts_10,
                total_ielts=total_ielts_10,
                total_vstep=total_vstep_10,
                timesheet_path=path_timesheet
            )
        except Exception as e:
            logging.error(f"Error generating Grade 10: {e}")
            return jsonify({'success': False, 'message': f'Lỗi tạo báo cáo Khối 10: {str(e)}'})

        try:
            out_dir_11 = os.path.join(app.config['GRADE_11_DIR'], f'Grade_11_Week {week_11}')
            generated_11, stats_11 = generate_grade_11_reports(
                week_11, vstep_11, ielts_11, path_11, output_dir=out_dir_11,
                target_classes_list=target_classes_11,
                course_vstep=course_vstep_11,
                course_ielts=course_ielts_11,
                total_ielts=total_ielts_11,
                total_vstep=total_vstep_11,
                timesheet_path=path_timesheet
            )
        except Exception as e:
            logging.error(f"Error generating Grade 11: {e}")
            return jsonify({'success': False, 'message': f'Lỗi tạo báo cáo Khối 11: {str(e)}'})

        response_data = {
            'success': True, 
            'message': f'Đã tạo thành công {len(generated_10) + len(generated_11)} báo cáo.',
            'reports_10': [os.path.basename(p) for p in generated_10],
            'reports_11': [os.path.basename(p) for p in generated_11],
            'week_10': week_10,
            'week_11': week_11,
            'stats_10': clean_nans(stats_10),
            'stats_11': clean_nans(stats_11)
        }
        
        # Log the response safely for debugging
        logging.info(f"Response Payload keys: {response_data.keys()}")
        
        return jsonify(response_data)

    except Exception as e:
        logging.error(f"General Error: {e}")
        return jsonify({'success': False, 'message': f'Lỗi hệ thống: {str(e)}'})

@app.route('/preview/<grade>/<week>/<filename>')
def preview_report(grade, week, filename):
    week_dir = f"{grade}_Week {week}"
    if grade == 'Grade_10':
        directory = os.path.join(app.config['GRADE_10_DIR'], week_dir)
    else:
        directory = os.path.join(app.config['GRADE_11_DIR'], week_dir)
        
    file_path = os.path.join(directory, filename)
    
    if not os.path.exists(file_path):
        return "File không tồn tại", 404
    
    try:
        with open(file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            return result.value
    except Exception as e:
        logging.error(f"Mammoth error: {e}")
        return f"Lỗi hiển thị nội dung: {str(e)}", 500


@app.route('/download/<grade>/<week>/<filename>')
def download_file(grade, week, filename):
    week_dir = f"{grade}_Week {week}"
    if grade == 'Grade_10':
        directory = os.path.join(app.config['GRADE_10_DIR'], week_dir)
    else:
        directory = os.path.join(app.config['GRADE_11_DIR'], week_dir)
        
    return send_from_directory(directory, filename, as_attachment=True)

@app.route('/download-zip/<grade>/<week>')
def download_zip(grade, week):
    week_dir_name = f"{grade}_Week {week}"
    if grade == 'Grade_10':
        base_dir = app.config['GRADE_10_DIR']
    else:
        base_dir = app.config['GRADE_11_DIR']
        
    target_dir = os.path.join(base_dir, week_dir_name)
    
    if not os.path.exists(target_dir):
        return jsonify({'success': False, 'message': 'Thư mục không tồn tại'})

    # Create Zip
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(target_dir):
            for file in files:
                zf.write(os.path.join(root, file), file)
    
    memory_file.seek(0)
    return send_file(memory_file, download_name=f'{week_dir_name}.zip', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
