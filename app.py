# app.py
from flask import Flask, render_template, request, jsonify
import os
from werkzeug.utils import secure_filename
from docx import Document
import io
import tempfile
import json
import uuid

# --- 配置 ---
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}  # 为简化，暂时只支持 docx。doc 支持需要额外库且复杂。
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
REGIONS_FILE = 'china_regions.json'

app = Flask(__name__)
app.secret_key = 'your_strong_secret_key_here'  # 更换为强密钥用于生产环境
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    """检查文件扩展名是否被允许"""
    try:
        app.logger.info(f"Checking file: {filename}")

        # 检查文件名是否为空
        if not filename:
            app.logger.warning("Empty filename provided")
            return False

        # 检查是否包含点号
        if '.' not in filename:
            app.logger.warning(f"No dot found in filename: {filename}")
            return False

        # 获取文件扩展名
        parts = filename.rsplit('.', 1)
        print(parts)
        app.logger.info(f"Filename parts: {parts}")

        if len(parts) != 2:
            app.logger.warning(f"Unexpected number of parts in filename: {filename}")
            return False

        extension = parts[1].lower()
        app.logger.info(f"Extension: {extension}, Allowed: {ALLOWED_EXTENSIONS}")

        result = extension in ALLOWED_EXTENSIONS
        app.logger.info(f"File allowed: {result}")
        return result
    except Exception as e:
        # 记录异常但不中断程序
        app.logger.error(f"Error checking file extension for filename '{filename}': {str(e)}")
        return False


def load_regions():
    """从 JSON 文件加载地域名称"""
    if os.path.exists(REGIONS_FILE):
        try:
            with open(REGIONS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 如果是列表格式，直接返回
                if isinstance(data, list):
                    return data
                # 如果是字典格式，提取所有地域名称
                elif isinstance(data, dict):
                    regions = []
                    # 遍历所有省份
                    for province, cities in data.items():
                        regions.append(province)
                        # 如果城市是字典形式（包含区县），提取所有城市和区县
                        if isinstance(cities, dict):
                            for city, districts in cities.items():
                                regions.append(city)
                                if isinstance(districts, list):
                                    regions.extend(districts)
                        # 如果城市是列表形式（直接列出城市），添加所有城市
                        elif isinstance(cities, list):
                            regions.extend(cities)
                    return regions
                else:
                    app.logger.error(f"JSON file {REGIONS_FILE} format is invalid.")
                    return []
        except (json.JSONDecodeError, IOError) as e:
            app.logger.error(f"Error reading {REGIONS_FILE}: {e}")
            return []

def flatten_regions(regions_data):
    """将嵌套的地域结构转换为扁平的名称列表"""
    regions = []
    if isinstance(regions_data, dict):
        for key, value in regions_data.items():
            regions.append(key)
            if isinstance(value, dict):
                regions.extend(flatten_regions(value))
            elif isinstance(value, list):
                regions.extend(value)
    elif isinstance(regions_data, list):
        regions.extend(regions_data)
    return regions

def save_regions(regions_data):
    """将地域名称保存到 JSON 文件"""
    try:
        with open(REGIONS_FILE, 'w', encoding='utf-8') as f:
            json.dump(regions_data, f, ensure_ascii=False, indent=4)
        return True
    except IOError as e:
        app.logger.error(f"Error saving to {REGIONS_FILE}: {e}")
        return False


class DocxTextExtractor:
    """专门用于提取 docx 文本和页码信息的类"""

    def __init__(self, docx_path):
        self.doc = Document(docx_path)
        self.page_breaks = self._find_page_breaks()
        self.full_text = self._extract_full_text()

    def _find_page_breaks(self):
        """估算段落在文档中的页码位置"""
        page_breaks = []
        current_page = 1
        current_line_count = 0
        lines_per_page_estimate = 40  # 这是一个估算值，可根据需要调整

        for i, paragraph in enumerate(self.doc.paragraphs):
            text = paragraph.text
            if text.strip():  # 忽略空段落
                # 估算段落行数
                lines_in_para = len(text) // 80 + text.count('\n') + 1  # 简单估算
                current_line_count += lines_in_para

                # 检查段落中是否有分页符
                if 'w:br' in paragraph._p.xml and 'type="page"' in paragraph._p.xml:
                    # 找到显式分页符
                    page_breaks.append((i, current_page))
                    current_page += 1
                    current_line_count = lines_in_para  # 重置行数计数

                # 检查是否因行数估算而翻页
                if current_line_count > lines_per_page_estimate:
                    page_breaks.append((i, current_page))
                    current_page += 1
                    current_line_count = lines_in_para  # 重置为当前段落的行数

        return page_breaks

    def _extract_full_text(self):
        """提取完整文本"""
        full_text = []
        for para in self.doc.paragraphs:
            # 保留换行符以便于上下文查找
            full_text.append(para.text + "\n")
        return ''.join(full_text)

    def get_page_number(self, paragraph_index):
        """根据段落索引估算页码"""
        page_num = 1
        for break_index, break_page in self.page_breaks:
            if paragraph_index >= break_index:
                page_num = break_page + 1  # 下一页开始
            else:
                break
        return page_num

    def find_keyword_occurrences(self, keywords):
        """查找关键词并返回其页码和上下文"""
        occurrences = []
        context_length = 50  # 上下文字符数

        # 为了提高查找效率，可以将列表转换为集合进行成员检查
        keywords_set = set(kw.strip() for kw in keywords if kw.strip())

        for kw_index, keyword in enumerate(keywords):
            keyword = keyword.strip()
            if not keyword:
                continue
            start = 0
            while True:
                pos = self.full_text.find(keyword, start)
                if pos == -1:
                    break

                # 找到包含关键词的段落索引
                para_index = 0
                text_pos = 0
                for i, para in enumerate(self.doc.paragraphs):
                    para_end = text_pos + len(para.text) + 1  # +1 for \n
                    if text_pos <= pos < para_end:
                        para_index = i
                        break
                    text_pos = para_end

                # 估算页码
                page_num = self.get_page_number(para_index)

                # 提取上下文
                context_start = max(0, pos - context_length)
                context_end = min(len(self.full_text), pos + len(keyword) + context_length)
                context = self.full_text[context_start:context_end].strip()

                occurrences.append({
                    'keyword': keyword,
                    'page': page_num,
                    'context': context
                })
                start = pos + 1  # 从下一个位置继续查找

        return occurrences


# 在app.py中添加或修改以下函数

def load_regions_structured():
    """从 JSON 文件加载结构化的地域名称"""
    if os.path.exists(REGIONS_FILE):
        try:
            with open(REGIONS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 确保返回的是字典格式
                if isinstance(data, dict):
                    return data
                else:
                    app.logger.error(f"JSON file {REGIONS_FILE} format is invalid (not a dict).")
                    return {"省级": [], "市级": [], "区级": []}
        except (json.JSONDecodeError, IOError) as e:
            app.logger.error(f"Error reading {REGIONS_FILE}: {e}")
            return {"省级": [], "市级": [], "区级": []}
    else:
        # 如果文件不存在，创建一个带有默认数据的文件
        default_regions = {
            "省级": ["北京", "天津", "上海", "重庆", "香港", "澳门", "内蒙古", "广西", "西藏", "宁夏", "新疆",
                     "河北", "山西", "辽宁", "吉林", "黑龙江", "江苏", "浙江", "安徽", "福建", "江西", "山东",
                     "河南", "湖北", "湖南", "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海"],
            "市级": ["石家庄", "唐山", "张家口", "太原", "沈阳", "大连", "辽阳", "长春", "松原", "延边",
                     "哈尔滨", "齐齐哈尔", "南京", "无锡", "徐州", "常州", "苏州", "南通", "连云港", "淮安",
                     "盐城", "扬州", "镇江", "泰州", "宿迁", "杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴",
                     "金华", "衢州", "舟山", "台州", "丽水", "合肥", "六安", "亳州", "福州", "厦门", "南昌",
                     "赣州", "济南", "青岛", "淄博", "郑州", "开封", "洛阳", "焦作", "武汉", "黄石", "十堰",
                     "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈", "咸宁", "随州", "恩施", "仙桃",
                     "潜江", "天门", "神农架林区", "长沙", "株洲", "湘潭", "岳阳", "广州", "深圳", "肇庆",
                     "惠州", "梅州", "海口", "三亚", "成都", "甘孜", "凉山", "贵阳", "黔西南", "黔东南", "黔南",
                     "昆明", "西安", "榆林", "兰州", "西宁"],
            "区级": ["东城", "西城", "朝阳", "丰台", "石景山", "海淀", "门头沟", "房山", "通州", "顺义",
                     "昌平", "大兴", "怀柔", "平谷", "密云", "延庆", "和平", "河东", "河西", "南开", "河北",
                     "红桥", "东丽", "西青", "津南", "北辰", "武清", "宝坻", "滨海新", "宁河", "静海", "蓟州",
                     "黄浦", "徐汇", "长宁", "静安", "普陀", "虹口", "杨浦", "闵行", "宝山", "嘉定", "浦东",
                     "金山", "松江", "青浦", "奉贤", "崇明", "万州", "涪陵", "渝中", "大渡口", "江北", "沙坪坝",
                     "九龙坡", "南岸", "北碚", "綦江", "大足", "渝北", "巴南", "黔江", "长寿", "江津", "合川",
                     "永川", "南川", "璧山", "铜梁", "潼南", "荣昌", "开州", "梁平", "武隆", "城口", "丰都",
                     "垫江", "忠县", "云阳", "奉节", "巫山", "巫溪"]
        }
        save_regions(default_regions)
        return default_regions


@app.route('/api/regions/structured', methods=['GET'])
def get_structured_regions():
    """获取结构化的地域名称"""
    regions = load_regions_structured()
    return jsonify({'success': True, 'regions': regions})


@app.route('/api/regions/<level>/<name>', methods=['DELETE'])
def delete_region_by_level(level, name):
    """根据级别删除地域名称"""
    data = request.get_json()
    password = data.get('password', '')

    # 验证密码
    if password != '123456':
        return jsonify({'success': False, 'message': '密码错误'}), 403

    regions = load_regions_structured()

    if level not in regions:
        return jsonify({'success': False, 'message': '无效的级别'}), 400

    if name not in regions[level]:
        return jsonify({'success': False, 'message': '地域名称不存在'}), 404

    regions[level].remove(name)

    if save_regions(regions):
        return jsonify({'success': True, 'message': '地域名称删除成功', 'name': name})
    else:
        return jsonify({'success': False, 'message': '保存地域名称失败'}), 500


@app.route('/api/regions/<level>', methods=['POST'])
def add_region_by_level(level):
    """根据级别添加地域名称"""
    regions = load_regions_structured()

    if level not in regions:
        return jsonify({'success': False, 'message': '无效的级别'}), 400

    data = request.get_json()
    new_region = data.get('name', '').strip()

    if not new_region:
        return jsonify({'success': False, 'message': '地域名称不能为空'}), 400

    if new_region in regions[level]:
        return jsonify({'success': False, 'message': '该地域名称已存在'}), 400

    regions[level].append(new_region)

    if save_regions(regions):
        return jsonify({'success': True, 'message': '地域名称添加成功', 'name': new_region})
    else:
        return jsonify({'success': False, 'message': '保存地域名称失败'}), 500


# --- 路由 ---

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    # 检查是否有文件部分
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有选择文件'}), 400
    file = request.files['file']

    # 如果用户没有选择文件
    if file.filename == '':
        return jsonify({'success': False, 'message': '没有选择文件'}), 400

    # 检查文件类型和保存
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        # 使用临时文件处理，避免磁盘写入
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1])
        file.save(temp_file.name)
        temp_file_path = temp_file.name
        temp_file.close()  # 关闭文件以便后续读取

        try:
            # 获取关键词类型
            check_type = request.form.get('checkType', 'custom')

            keywords = []
            if check_type == 'custom':
                keywords_str = request.form.get('keywords', '')
                if not keywords_str.strip():
                    return jsonify({'success': False, 'message': '请输入至少一个关键词'}), 400
                keywords = keywords_str.split('，')  # 使用中文逗号分割
            elif check_type == 'china_regions':
                keywords = load_regions()  # 从 JSON 文件加载地域名
                if not keywords:
                    return jsonify({'success': False, 'message': '地域名称列表为空或加载失败'}), 500
            else:
                return jsonify({'success': False, 'message': '无效的检查类型'}), 400

            # 提取文本和查找关键词
            if filename.lower().endswith('.docx'):
                extractor = DocxTextExtractor(temp_file_path)
                occurrences = extractor.find_keyword_occurrences(keywords)
                os.unlink(temp_file_path)  # 处理完立即删除临时文件
                return jsonify(
                    {'success': True, 'filename': filename, 'occurrences': occurrences, 'checkType': check_type})

            else:
                os.unlink(temp_file_path)
                return jsonify({'success': False, 'message': '不支持的文件类型。请上传 .docx 文件。'}), 400

        except Exception as e:
            # 确保即使出错也删除临时文件
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
            app.logger.error(f"Error processing file {filename}: {e}")
            return jsonify({'success': False, 'message': f'处理文件时出错: {str(e)}'}), 500

    else:
        return jsonify({'success': False, 'message': '不支持的文件类型。请上传 .docx 文件。'}), 400


# --- 地域名称管理 API ---

@app.route('/api/regions', methods=['GET'])
def get_regions():
    """获取所有地域名称"""
    regions = load_regions()
    return jsonify({'success': True, 'regions': regions})


@app.route('/api/regions', methods=['POST'])
def add_region():
    """添加一个新的地域名称"""
    data = request.get_json()
    new_region = data.get('region', '').strip()

    if not new_region:
        return jsonify({'success': False, 'message': '地域名称不能为空'}), 400

    regions = load_regions()
    if new_region in regions:
        return jsonify({'success': False, 'message': '该地域名称已存在'}), 400

    regions.append(new_region)
    if save_regions(regions):
        return jsonify({'success': True, 'message': '地域名称添加成功', 'region': new_region})
    else:
        return jsonify({'success': False, 'message': '保存地域名称失败'}), 500


@app.route('/api/regions/<region>', methods=['PUT'])
def update_region(region):
    """修改一个地域名称"""
    data = request.get_json()
    updated_region = data.get('updatedRegion', '').strip()

    if not updated_region:
        return jsonify({'success': False, 'message': '新地域名称不能为空'}), 400

    regions = load_regions()
    if region not in regions:
        return jsonify({'success': False, 'message': '要修改的地域名称不存在'}), 404

    if updated_region in regions:
        return jsonify({'success': False, 'message': '新地域名称已存在'}), 400

    index = regions.index(region)
    regions[index] = updated_region
    if save_regions(regions):
        return jsonify(
            {'success': True, 'message': '地域名称修改成功', 'oldRegion': region, 'newRegion': updated_region})
    else:
        return jsonify({'success': False, 'message': '保存地域名称失败'}), 500


@app.route('/api/regions/<region>', methods=['DELETE'])
def delete_region(region):
    """删除一个地域名称"""
    regions = load_regions()
    if region not in regions:
        return jsonify({'success': False, 'message': '要删除的地域名称不存在'}), 404

    regions.remove(region)
    if save_regions(regions):
        return jsonify({'success': True, 'message': '地域名称删除成功', 'region': region})
    else:
        return jsonify({'success': False, 'message': '保存地域名称失败'}), 500


if __name__ == '__main__':
    app.run(debug=True)  # 生产环境中请设置 debug=False