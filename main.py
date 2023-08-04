import pytesseract as pts
import numpy as np
import os
from docx import Document
import sys
import cv2
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import matplotlib.pyplot as plt
import func_dewarping as fdw
DEBUG_LEVEL = 0 
def save_string_to_docx():
    content = text_box.get("1.0", "end-1c")  # Lấy nội dung từ hộp văn bản
    file_path = fd.asksaveasfilename(defaultextension=".docx",
                                             filetypes=[("Microsoft Word Documents", "*.docx")])
    if file_path:
        doc = Document()
        doc.add_paragraph(content)
        doc.save(file_path)
        status_label.config(text=f"File đã được lưu thành công tại: {file_path}")
    else:
        status_label.config(text="Không lưu file.")
def process_string(input_string):
    lines = input_string.split('\n')
    result_lines = ''''''
    last_char = ''
    for line in lines:
        if line.strip():
            first_char = line.lstrip()[0]
            print(first_char.islower())
            if len(line) == 1 and line.lstrip()[0] == '\n':
                last_char = '.'
                break
            if (first_char.islower() and last_char != '.'):
                result_lines = result_lines + ' ' + line
            else:
                if ((first_char.islower() == False and last_char == '.') or first_char == '-'):
                    result_lines = result_lines + '\n' + line
                else: result_lines = result_lines + ' ' + line
            last_char = line.lstrip()[len(line)-1]
            print(last_char)

    return result_lines
root = tk.Tk()
root.title("Chương trình lưu chuỗi vào file docx")
pts.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
file_paths = fd.askopenfilenames(filetypes=[('Image Files', ['.jpeg', '.jpg', '.png'])])
for file_path in file_paths:
    # gray = cv2.imread(file_path, cv2.IMREAD_GRAYSCALE)
    # histogram = cv2.calcHist([gray], [0], None, [256], [0, 256])
    # threshold_value = int(np.mean(np.where(histogram == np.max(histogram))))
    img = cv2.imread(file_path)
    small = fdw.resize_to_screen(img)
    basename = os.path.basename(file_path)
    name, _ = os.path.splitext(basename)
    print ('loaded', basename, 'with size', fdw.imgsize(img),)
    print ('and resized to', fdw.imgsize(small))

    if DEBUG_LEVEL >= 3:
        fdw.debug_show(name, 0.0, 'original', small)

    pagemask, page_outline = fdw.get_page_extents(small)

    cinfo_list = fdw.get_contours(name, small, pagemask, 'text')
    spans = fdw.assemble_spans(name, small, pagemask, cinfo_list)

    if len(spans) < 3:
        print ('  detecting lines because only', len(spans), 'text spans')
        cinfo_list = fdw.get_contours(name, small, pagemask, 'line')
        spans2 = fdw.assemble_spans(name, small, pagemask, cinfo_list)
        if len(spans2) > len(spans):
            spans = spans2

    if len(spans) < 1:
        print( 'skipping', name, 'because only', len(spans), 'spans')
        continue

    span_points = fdw.sample_spans(small.shape, spans)

    print( '  got', len(spans), 'spans',)
    print ('with', sum([len(pts) for pts in span_points]), 'points.')

    corners, ycoords, xcoords = fdw.keypoints_from_samples(name, small,
                                                           pagemask,
                                                           page_outline,
                                                           span_points)

    rough_dims, span_counts, params = fdw.get_default_params(corners,
                                                             ycoords, xcoords)

    dstpoints = np.vstack((corners[0].reshape((1, 1, 2)),) +
                              tuple(span_points))

    params = fdw.optimize_params(name, small,
                                 dstpoints,
                                 span_counts, params)

    page_dims = fdw.get_page_dims(corners, rough_dims, params)

    remap_image,threshold = fdw.remap_image(name, img, small, page_dims, params)
    data_raw = pts.image_to_string(threshold, lang='vie')
    print('data_raw = ', data_raw)
    plt.subplot(1, 3, 1)
    plt.imshow(img)
    plt.axis('off')
    plt.title('Ảnh gốc')
    plt.subplot(1, 3, 2)
    plt.imshow(remap_image, cmap='gray')
    plt.title('Ảnh sau khi Dewarp')
    plt.subplot(1, 3, 3)
    plt.imshow(threshold, cmap='gray')
    plt.title('Ảnh sau khi xét ngưỡng')
    plt.show()
content = process_string(data_raw)
text_box = tk.Text(root, wrap=tk.WORD, width=40, height=10)
text_box.pack(pady=10)
text_box.delete("1.0", tk.END)  # Xóa nội dung hiện tại của hộp văn bản
text_box.insert("1.0", content)  # Thêm nội dung mới vào hộp văn bản
# Tạo nút lưu file
save_button = tk.Button(root, text="Lưu vào file docx", command=save_string_to_docx)
save_button.pack()
# Tạo nhãn để hiển thị trạng thái
status_label = tk.Label(root, text="")
status_label.pack()
# Bắt đầu vòng lặp chạy chương trình
root.mainloop()