# PDF2PPT-ReportConverter
By GPT-4o & TJJ

使用 PyMuPDF 和 pptx 将医疗检查报告转为 PPT

因为涉及个人信息，原版 PDF 无法放到仓库中

---

### 代码结构概述

代码整体分为以下几个模块：

1. **`extract_info_from_page(page)`** - 文本信息提取模块
2. **`process_image(page)`** - 图像处理与裁剪模块
3. **`add_slide_with_image(presentation, cropped_img, check_info, date_info)`** - PPT 幻灯片生成模块
4. **`process_pdf_to_ppt(pdf_path, output_ppt_path)`** - 主处理流程模块
5. **`main()`** - 主程序

每个模块的功能如下：

---

### 1. `extract_info_from_page(page)`

**功能：**

- 从指定的 PDF 页面中提取文本信息，并利用正则表达式提取特定的日期和检查项目信息。

**关键步骤：**

- 利用正则表达式 `date_pattern` 提取日期信息。
- 利用正则表达式 `check_item_pattern` 提取检查项目信息。
- 返回提取到的日期和检查项目。

**逻辑：**

```python
# Define the regular expression patterns for extracting date and check items
date_pattern = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")
check_item_pattern = re.compile(r"检验项目[:：]\s*(.*?)(?:\n)")

# Using regular expressions to find '采集时间' and '检验项目'
date_match = date_pattern.search(text)
check_info_match = check_item_pattern.search(text)

date_info = date_match.group(0)[:10] if date_match else "采集时间未找到"
check_info = check_info_match.group(1) if check_info_match else "检验项目未找到"
```

对于日期，匹配文本中遇到的第一个 "yyyy-mm-dd hh:mm" 格式，保留年月日

对于检验项目，匹配文本中遇到的第一个 "检验项目: xxx\n"

---

### 2. `process_image(page)`

**功能：**

- 将 PDF 页面转换为高分辨率图像，并对图像进行裁剪，去除页面中的多余空白区域，保留主要内容。

**处理流程：**

1. **初步裁剪：**
   - 对页面图像进行初步裁剪，移除顶部 16% 和底部 5% 的部分，保留主要内容区域。

2. **灰度转换：**
   - 将图像转换为灰度图像，以便更容易检测像素的亮度。

3. **从上往下扫描：**
   - 扫描图像的每一行，寻找黑色像素比例超过 60% 的第一行，并记录为 `upper_bound`。这一行将作为图像内容的上边界。

4. **从下往上扫描并定位底部裁剪点：**
   - 首先从底部开始向上扫描，找到黑色像素比例超过 60% 的第一行，并记为 `lower_black_bound`。
   - 从 `lower_black_bound` 上方两行开始，继续向上扫描，找到第一个非全白的行，记为 `lower_bound`。这一行将作为图像内容的下边界。

5. **最终裁剪：**
   - 使用 `upper_bound` 和 `lower_bound` 对图像进行最终裁剪，移除顶部和底部的空白区域。

**关键参数：**

- **DPI (dots per inch)**：控制图像的分辨率，建议设为 300 以确保图像清晰度。
- **`threshold`**：用于定义一行中黑色像素的比例，通常设置为 153（约 60%）。

---

### 3. `add_slide_with_image(presentation, cropped_img, check_info, date_info)`

**功能：**

- 创建一个新的 PPT 幻灯片，并将处理后的图像和文本信息添加到幻灯片中。

**处理流程：**

1. **创建幻灯片：**
   - 创建一个空白布局的幻灯片。

2. **添加标题：**
   - 在幻灯片左上角添加一个标题文本框，并设置字体、大小、加粗等样式。

3. **动态调整图像大小与位置：**
   - 根据图像的宽高比和幻灯片的尺寸，动态调整图像的大小，使其在幻灯片中居中显示并占用合适的空间。

4. **插入图像：**
   - 将裁剪后的图像插入到幻灯片中，位置与大小由前一步计算得出。

**关键参数：**

- **幻灯片宽度 (`slide_width`)**：默认设置为 10 英寸。
- **幻灯片高度 (`slide_height`)**：默认设置为 6 英寸，留出顶部标题部分。

---

### 4. `process_pdf_to_ppt(pdf_path, output_ppt_path)`

**功能：**

- 该函数是主处理流程模块，用于遍历整个 PDF 文件的每一页，并依次调用上述函数进行处理，最终生成包含所有处理内容的 PPT 文件。

**处理流程：**

1. **加载 PDF：**
   - 使用 `fitz` 库加载指定路径的 PDF 文件。

2. **逐页处理：**
   - 对 PDF 的每一页进行处理：
     - 提取文本信息。
     - 处理图像并裁剪空白。
     - 将处理后的内容添加到 PPT 幻灯片中。

3. **保存 PPT 文件：**
   - 将处理结果保存为一个新的 PPT 文件，路径由 `output_ppt_path` 指定。

---

### 5. `main()`

**参数：**

1. **pdf_path：**
   - 输入 PDF 路径，默认为 `检查报告.pdf`

2. **ppt_path：**
   - 输出 PPT 路径，默认为 `检查报告.ppt`

---

### 参数调整

- **DPI (dots per inch)**：在 `process_image` 函数中，`dpi=300` 适用于大多数高质量图像处理需求。如果需要更高的图像清晰度，可以适当提高此值，但可能导致处理时间和内存使用增加。

- **裁剪比例**：初步裁剪中，顶部 17.2% 和底部 11% 的移除比例根据实际需求可以进行调整。如果页面内容较多或较少，可以适当调整这些比例以适应不同的文档格式。

- **`threshold` 值**：决定了像素的“黑色”标准，153 是一个大致的参考值。如果图像中的颜色较淡，可以降低此值，以确保可以准确识别“黑色”区域。

---

### 使用

安装依赖

```python
pip install PyMuPDF Pillow python-pptx
```

代码使用

- 脚本运行

```python
python report_to_ppt.py # 默认输入和输出
python report_to_ppt.py 检查报告.pdf #默认输出
python report_to_ppt.py 检查报告.pdf 检查报告.pptx #相对路径
python report_to_ppt.py D:\Projects\Jupyter\hospital_report\检查报告_full.pdf D:\Projects\Jupyter\hospital_report\检查报告.pptx #绝对路径
```

- exe 运行

```python
 pyinstaller --onefile .\report_to_ppt.py #导出为exe
```

```python
report_to_ppt.exe # 默认输入和输出
report_to_ppt.exe 检查报告.pdf #默认输出
report_to_ppt.exe 检查报告.pdf 检查报告.pptx #相对路径
report_to_ppt.exe D:\Projects\Jupyter\hospital_report\检查报告_full.pdf D:\Projects\Jupyter\hospital_report\检查报告.pptx #绝对路径
```

- 文件目录版

遍历当前目录下的报告 PDF，生成同名 PPTX

对应 report_to_ppt_auto.py
