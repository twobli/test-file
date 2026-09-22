from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from pathlib import Path

OUT = Path(r"C:\Users\Windows\Documents\Codex\2026-09-22\new-chat\outputs\基于基因的眼病慢病诊断应用需求规格说明_最新版.docx")

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ['val','sz','space','color']:
                if key in kwargs[edge]:
                    element.set(qn('w:{}'.format(key)), str(kwargs[edge][key]))

def set_cell_text(cell, text, bold=False, color=None, size=9.5):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(str(text))
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = 'Microsoft YaHei'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def add_table(doc, headers, rows, widths=None, font_size=9):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr = table.rows[0]
    hdr._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    for i, h in enumerate(headers):
        set_cell_text(hdr.cells[i], h, bold=True, color='FFFFFF', size=font_size)
        set_cell_shading(hdr.cells[i], '1F4E79')
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            set_cell_text(cells[i], val, size=font_size)
            if len(table.rows) % 2 == 0:
                set_cell_shading(cells[i], 'F4F7FA')
    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={'val':'single','sz':'4','color':'D9E2F3'}, bottom={'val':'single','sz':'4','color':'D9E2F3'}, left={'val':'single','sz':'4','color':'D9E2F3'}, right={'val':'single','sz':'4','color':'D9E2F3'})
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table

def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.12
        r = p.add_run(item)
        r.font.size = Pt(10.5)
        r.font.name = 'Microsoft YaHei'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

def add_numbered(doc, items):
    for item in items:
        p = doc.add_paragraph(style='List Number')
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.12
        r = p.add_run(item)
        r.font.size = Pt(10.5)
        r.font.name = 'Microsoft YaHei'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

def add_para(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.18
    if bold_prefix and text.startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        r1.bold = True
        r1.font.name = 'Microsoft YaHei'; r1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei'); r1.font.size = Pt(10.5)
        r2 = p.add_run(text[len(bold_prefix):])
        r2.font.name = 'Microsoft YaHei'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei'); r2.font.size = Pt(10.5)
    else:
        r = p.add_run(text)
        r.font.name = 'Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei'); r.font.size = Pt(10.5)
    return p

def add_callout(doc, label, text, fill='FFF2CC', color='7F6000'):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.cell(0,0)
    set_cell_shading(cell, fill)
    set_cell_border(cell, top={'val':'single','sz':'8','color':'D6B656'}, bottom={'val':'single','sz':'8','color':'D6B656'}, left={'val':'single','sz':'8','color':'D6B656'}, right={'val':'single','sz':'8','color':'D6B656'})
    cell.text=''
    p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.12
    r=p.add_run(label+'：'); r.bold=True; r.font.color.rgb=RGBColor.from_string(color); r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft YaHei'); r.font.size=Pt(10)
    r=p.add_run(text); r.font.color.rgb=RGBColor.from_string(color); r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft YaHei'); r.font.size=Pt(10)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)

# Document setup
doc = Document()
sec = doc.sections[0]
sec.page_width = Inches(8.5)
sec.page_height = Inches(11)
sec.top_margin = Inches(0.65)
sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.75)
sec.right_margin = Inches(0.75)

styles = doc.styles
for name in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
    st = styles[name]
    st.font.name = 'Microsoft YaHei'
    st._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
styles['Normal'].font.size = Pt(10.5)
styles['Normal'].paragraph_format.space_after = Pt(5)
styles['Normal'].paragraph_format.line_spacing = 1.18
styles['Title'].font.size = Pt(22); styles['Title'].font.bold=True; styles['Title'].font.color.rgb=RGBColor(31,78,121)
styles['Heading 1'].font.size = Pt(15); styles['Heading 1'].font.bold=True; styles['Heading 1'].font.color.rgb=RGBColor(31,78,121)
styles['Heading 1'].paragraph_format.space_before=Pt(12); styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].font.size = Pt(12.5); styles['Heading 2'].font.bold=True; styles['Heading 2'].font.color.rgb=RGBColor(47,84,150)
styles['Heading 2'].paragraph_format.space_before=Pt(9); styles['Heading 2'].paragraph_format.space_after=Pt(4)
styles['Heading 3'].font.size = Pt(11); styles['Heading 3'].font.bold=True; styles['Heading 3'].font.color.rgb=RGBColor(68,68,68)

# Header/footer
header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hr = header.add_run('基于基因的眼病/慢病诊断应用｜需求规格说明｜最新版')
hr.font.size = Pt(8); hr.font.color.rgb = RGBColor(128,128,128); hr.font.name='Microsoft YaHei'; hr._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft YaHei')
footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = footer.add_run('内部产品与研发评审稿  |  以临床专业人员审核为前提')
fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(128,128,128); fr.font.name='Microsoft YaHei'; fr._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft YaHei')

# Cover
title = doc.add_paragraph(style='Title')
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.add_run('基于基因的眼病慢病诊断应用需求规格说明')
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r=sub.add_run('基于现有 HTML 原型的产品叙事与可研发可测试规格')
r.font.size=Pt(12); r.font.color.rgb=RGBColor(89,89,89); r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft YaHei')
doc.add_paragraph()
add_table(doc, ['文档属性','内容'], [
    ('文档范围','仅覆盖“基于基因的眼病/慢病诊断应用”这一产品，不包含其他需求文档'),
    ('材料依据','最新版原型：基于基因的眼病_慢病诊断应用.html（文件更新时间 2026-09-20）；原需求文档：基于基因的眼病慢病诊断应用需求规格说明.docx；结合 requirement-refiner skill 梳理'),
    ('版本信息','V2.0｜根据 2026-09-20 原型更新，生成日期 2026-09-22'),
    ('文档状态','更新版需求基线草案，供产品、临床、研发、算法、测试和数据安全评审'),
    ('目标读者','产品经理、临床专家、算法/数据团队、前后端研发、测试与运维'),
    ('重要边界','系统定位为临床辅助与科研分析工具；任何风险提示不能替代临床诊断、遗传咨询或治疗决策'),
], widths=[1.35,5.7], font_size=9.5)
doc.add_page_break()

# 1 Overview
doc.add_heading('1 文档说明', level=1)
add_para(doc, '本需求规格说明用于定义基于基因、眼底影像和临床背景信息的眼病/慢病诊断辅助应用。文档先描述业务问题、角色、目标和端到端流程，再给出页面、字段、交互、状态、权限、数据、异常和验收规则。')
add_para(doc, '本次更新以最新版原型已呈现的三步链路为基线：输入信息、智能分析、评估报告。相较原文档，本版补充统一多选基因分型面板及已选摘要、CFP/OCT/OCTA与单眼/双眼影像配置、四阶段分析状态、基因变异解读小结、2型糖尿病/DR/UM三类综合风险展示，以及BAP1、SF3B1、EIF1AX和染色体3的冲突提示。原型中的评分、病灶标注、方向性建议、计时动画和导出实现仍以演示或前端占位为主，生产化规则必须经过临床、算法、产品和合规评审后才能作为正式能力上线。')
add_callout(doc, '事实与假设边界', '原型明确展示的页面、字段和操作作为已确认的产品意图；原型未说明的角色、接口、数据保留周期、审计和指标在本文中标为工作假设或待产品确认，不把通用经验默认为既定业务规则。', fill='EAF2F8', color='1F4E79')

# 2 Problem goals
doc.add_heading('2 业务问题与产品目标', level=1)
doc.add_heading('2.1 业务问题', level=2)
add_bullets(doc, [
    '眼科相关基因检测结果、眼底影像和患者慢病背景通常分散在不同材料中，医护人员需要手工汇总，影响初筛、会诊和科研分析效率。',
    '不同疾病方向的基因证据强度、影像表型和临床提示需要在同一份结果中分层展示，避免把研究阶段证据误认为确定性结论。',
    '现有原型希望通过影像先判别疾病方向，再将基因型证据与影像异常共同纳入综合评估，形成结构化报告。'
])
doc.add_heading('2.2 产品目标', level=2)
add_bullets(doc, [
    '支持用户录入患者可选基本信息、选择基因/分子标志物结果并上传眼底影像。',
    '通过统一的分析流程输出影像异常提示、基因变异解读、疾病风险分层和方向性建议。',
    '以证据强度、数据来源和免责声明支撑临床辅助、MDT 会诊和脱敏科研分析，不输出未经审核的确定性诊断或治疗处方。',
    '为后续算法服务、报告导出、数据审计和临床系统对接预留清晰的数据对象与接口边界。'
])
doc.add_heading('2.3 成功标准', level=2)
add_table(doc, ['维度','可观察结果','当前状态'], [
    ('流程完成','用户能够从输入页完成信息录入、影像上传并进入分析和报告页','原型已展示；生产交互待实现'),
    ('结果完整','报告至少包含患者背景、基因易感性、影像提示、基因解读、综合风险和方向性建议','原型已展示'),
    ('证据可解释','每个基因或标志物可看到证据强度、适用方向和必要的限制说明','原型部分展示；需产品化'),
    ('安全可控','系统在关键位置明确“仅供参考/不能替代临床诊断”，并保留版本与审计信息','原型已展示免责声明；审计待确认'),
    ('可对接','报告和分析结果具备可被算法接口或业务系统调用的结构化数据','原型仅有接口文档意图；接口契约待确认'),
], widths=[1.1,3.7,2.25])

# 3 scope
doc.add_heading('3 范围与边界', level=1)
doc.add_heading('3.1 本期范围', level=2)
add_bullets(doc, [
    '患者基本信息录入：姓名、年龄、性别、既往病史，均允许选填。',
    '基因分型/分子标志物多选录入：糖尿病及糖尿病视网膜病变相关位点，以及脉络膜黑色素瘤相关驱动和预后标志物。',
    '眼底影像输入：眼底彩照（CFP）、OCT、OCTA；支持单眼或双眼配置；支持本地图像上传。',
    '三步流程：输入信息、智能分析、评估报告。',
    '报告模块：患者背景、关键疾病基因易感性评分、影像异常提示与特征说明、基因变异解读、基因变异解读小结、影像与基因综合疾病风险评估（同时展示2型糖尿病、DR和UM方向）、方向性建议。',
    '报告与接口文档导出：当前原型提供“导出综合评估报告”（PDF）和“导出算法接口文档”（.doc）两个入口；当前未展示 Word 版综合评估报告按钮，是否保留为后续能力需另行确认。',
])
doc.add_heading('3.2 明确不在本期范围', level=2)
add_bullets(doc, [
    '不直接作出具有法律或临床效力的疾病确诊、分期、治疗处方或用药调整。',
    '不替代遗传咨询，不基于多基因风险提示直接对无症状亲属下达预测性检测结论。',
    '不在本期自行扩展到原型未覆盖的疾病、基因、影像类型或治疗方案。',
    '不把原型中的示意评分阈值、演示病灶、默认疾病路径和模拟分析时长直接视为临床验证规则。',
    '不在本文定义具体数据库表结构、模型网络结构、部署架构或第三方厂商选型。'
])
doc.add_heading('3.3 关键待确认项', level=2)
add_table(doc, ['优先级','待确认项','影响'], [
    ('P0','产品是否只服务医护/科研人员，是否允许患者端直接查看报告','决定登录、隐私、文案、权限和免责声明'),
    ('P0','未勾选基因项目到底表示“未检测”“阴性”还是“未提供”','影响风险计算、报告措辞和临床可信度'),
    ('P0','正式算法服务的输入输出、模型版本、置信度和失败重试协议','影响后端接口、状态机和验收'),
    ('P0','正式支持的疾病清单、基因清单、证据等级和版本更新流程','影响知识库、报告可追溯和临床审核'),
    ('P1','患者身份标识、影像/基因数据存储时长、删除和脱敏策略','影响合规、权限和运维'),
    ('P1','PDF 综合评估报告的模板、签名、机构信息和报告编号规则','影响 PDF 导出实现及临床使用'),
    ('P1','算法接口文档 .doc 的字段冻结、示例、错误码和版本发布规则','影响接口对接与研发测试；当前原型为浏览器端离线生成'),
    ('P1','Word 版综合评估报告是否作为后续能力保留','当前最新版原型未展示该入口，不应作为本版已实现功能验收'),
], widths=[0.65,3.45,2.95])

# 4 roles
doc.add_heading('4 用户角色与权限', level=1)
add_para(doc, '以下角色为支持实现和测试的工作假设；若项目已有统一身份与权限体系，应以其角色编码为准。最小权限原则要求：能查看结果不等于能修改知识库、下载原始基因/影像数据或管理系统配置。')
add_table(doc, ['角色','主要任务','可见数据','操作权限'], [
    ('临床医生/眼科医生','录入或核对患者资料，发起分析，查看和导出报告','本人有权限的患者基本信息、影像、基因结果和报告','创建分析、查看报告、重新分析、导出；不能修改已发布的证据知识库'),
    ('遗传/检验专业人员','核对基因检测结果及证据解释','授权患者基因结果、证据来源与版本信息','查看、补充或复核基因数据；是否可编辑待确认'),
    ('科研用户','进行脱敏病例分析和结果导出','仅可见脱敏数据及被授权的报告字段','发起科研分析、查看脱敏报告；不可见直接身份信息'),
    ('系统管理员','维护用户、权限、字典、接口和运行状态','按运维授权范围访问，不默认读取全部医疗内容','配置管理、审计查询、失败任务处置，不参与临床结论修改'),
], widths=[1.25,2.25,2.4,2.0], font_size=8.7)
add_callout(doc, '权限边界', '无权限用户不得通过页面、导出、接口或猜测对象标识访问患者资料、原始影像、基因结果和报告；权限校验必须在前端展示控制之外由服务端执行。', fill='FCE4D6', color='9E480E')

# 5 overall process
doc.add_heading('5 端到端业务流程', level=1)
add_para(doc, '用户从输入页开始，至少提供一种眼底影像后发起一键分析。系统进入分析态，按步骤展示多模态分析进度；分析成功后进入综合评估报告。报告既是本次任务的终点，也是后续会诊、科研和导出的数据来源。')
add_para(doc, '主流程：患者/病例选择 → 输入基本信息 → 选择基因/分子标志物 → 配置眼底影像类型和眼数 → 上传至少一张影像 → 发起分析 → 影像疾病方向判别 → 基因、影像与临床证据融合 → 生成结构化报告 → 查看/导出/重新分析。')
add_table(doc, ['阶段','触发与前置条件','关键动作','输出与状态'], [
    ('输入信息','用户已进入应用，具备病例或患者授权；基本信息可为空','录入患者信息；多选基因/标志物；配置 CFP/OCT/OCTA 与单眼/双眼；上传图像','形成一次分析草稿，状态为草稿'),
    ('校验','用户点击开始分析','检查至少一张有效影像、文件类型/大小/可读性、标志物互斥项和必填上下文','校验通过进入分析中；失败留在输入页并定位错误'),
    ('智能分析','输入校验通过且任务已创建','执行影像判别、特征提取、证据关联、综合风险评估','任务状态按阶段变化：排队/处理中/成功/失败/取消'),
    ('评估报告','分析成功且结果完整或带可解释的部分结果','展示患者背景、基因、影像、风险和建议；允许导出或重新分析','生成报告版本，记录数据与模型版本'),
    ('异常终止','上传失败、算法失败、超时、权限失效或用户取消','展示原因、可恢复动作和任务编号；不生成伪造结论','任务为失败/取消，原输入草稿按策略保留或清理'),
], widths=[1.0,2.15,2.85,2.0], font_size=8.8)

# 6 page map
doc.add_heading('6 页面地图与跨页面传递', level=1)
add_table(doc, ['页面/状态','入口','主要内容','离开页面时保留'], [
    ('输入信息','应用首页或新建分析','患者信息、基因/标志物、影像配置和上传控件','病例标识、输入草稿、上传对象引用、知识库版本'),
    ('智能分析','输入校验通过后自动进入','动画/进度、分析步骤、当前提示','分析任务 ID、步骤状态、失败原因、取消状态'),
    ('评估报告','分析成功后自动进入；也可从历史任务进入','结构化报告及免责声明','报告 ID、报告版本、生成时间、模型/知识库版本、导出状态'),
    ('重新分析','报告页操作','返回输入页，默认带回上一次可复用输入','是否复用原影像与基因结果由产品确认；不得无提示覆盖原报告'),
], widths=[1.2,1.55,3.2,2.05], font_size=8.8)
add_para(doc, '刷新与返回规则：输入页刷新应尽量恢复未提交草稿；分析中刷新应通过任务 ID 查询服务端状态；报告页刷新应通过报告 ID重新读取，不应依赖浏览器内存中的演示数据。若数据已过期、被删除或权限发生变化，页面应提示并停止展示敏感内容。')

# 7 input page
doc.add_heading('7 模块需求 输入信息', level=1)
add_heading = doc.add_heading
add_heading('7.1 页面定位与使用过程', level=2)
add_para(doc, '输入信息页用于收集一次分析所需的患者背景、基因/分子标志物结果和眼底影像。基本信息是可选的；影像是进入分析的最低输入要求；基因结果用于补充遗传或分子证据，不应因未提供基因结果而阻断仅基于影像的辅助分析，具体是否允许需产品确认。')
add_para(doc, '页面初始展示三个卡片区域：患者基本信息、基因分型录入、眼底影像上传，以及“开始一键分析”按钮。用户完成任意可用信息输入后，系统应即时显示已选择项目、上传预览和校验状态；点击开始分析后执行前置校验。')
add_heading('7.2 患者基本信息字段', level=2)
add_table(doc, ['字段','类型/控件','是否必填','校验与展示规则','报告映射'], [
    ('姓名','单行文本','否','长度、非法字符和敏感信息策略待确认；为空显示未提供','患者背景信息'),
    ('年龄','整数输入','否','原型限制 0–120；生产规则需确认是否允许未知/区间年龄','患者背景信息'),
    ('性别','单选下拉','否','原型选项：选填、男、女；是否增加未知/其他待确认','患者背景信息'),
    ('既往病史','单行文本/多行文本','否','支持慢病、病程等自由描述；生产版建议拆分结构化病史与备注','患者背景信息和分析上下文'),
], widths=[1.0,1.45,0.75,3.0,1.5], font_size=8.7)
add_heading('7.3 基因分型与分子标志物', level=2)
add_para(doc, '最新版原型采用统一的多选下拉面板，将2型糖尿病易感位点、糖尿病视网膜病变候选基因、脉络膜黑色素瘤驱动突变及预后相关标志物放在同一录入入口；页面下方新增“已选择的基因分型”标签摘要。BAP1阳性/阴性、SF3B1突变型/野生型、EIF1AX突变型/野生型、染色体3单体/二体为互斥选项；驱动突变按单项选择处理。原型提示未勾选项目按“未检测/阴性”处理，但生产系统不得继续混用这两个语义：至少应在数据模型中区分“未提供、未检测、阴性、阳性、异常、质量不合格”，并在报告中按真实状态表述。')
add_table(doc, ['分组','原型覆盖项目','输入规则','证据/互斥规则'], [
    ('2型糖尿病易感性','TCF7L2、PPARG、KCNJ11、SLC30A8、CDKAL1、IGF2BP2','多选；显示基因名、位点/序列摘要和说明','证据强度按强/中等/较弱展示；不得将示意分值当作临床风险概率'),
    ('糖尿病视网膜病变候选位点','VEGFA、AKR1B1、ACE、AGER、NOS3、APOE','多选；展示适用方向和证据限制','较弱或结论不一致的项目只作背景提示，不单独驱动高风险结论'),
    ('脉络膜黑色素瘤驱动突变','GNAQ、GNA11、CYSLTR2、PLCB4','统一多选面板中的驱动突变单项选择；驱动事件属于肿瘤起始事件，不作为预后指标；未检出/未测由生产状态模型表达','同组驱动突变不可同时选择；未输入不得静默解释为未检出'),
    ('预后相关标志物','BAP1 阳性/阴性；SF3B1 突变型/野生型；EIF1AX 突变型/野生型；染色体3单体/二体','同一标志物的相反状态互斥；未测应单独表达','BAP1与染色体3不一致时提示复核；BAP1阴性同时伴SF3B1/EIF1AX突变时提示少见共存组合，不按最高风险直接判定'),
], widths=[1.45,2.55,2.0,1.95], font_size=8.1)
add_heading('7.4 眼底影像上传', level=2)
add_table(doc, ['字段/操作','规则','错误与反馈'], [
    ('影像类型','原型支持眼底彩照 CFP、OCT、OCTA；一次配置一种类型','类型变更后同步影像槽位；若原影像不再适用，提示清理或重新上传'),
    ('上传眼数','原型支持单眼或双眼；双眼生成右眼/左眼独立槽位','双眼配置下右眼/左眼分别管理；是否允许只上传一眼由产品确认，不能把一张图静默映射到双眼'),
    ('本地图像上传','支持本地图像上传，前端文件控件接受 image/*；生产需定义格式、大小、分辨率、方向和去元数据规则','格式不支持、读取失败、过大、病毒扫描失败时阻止进入分析并给出可恢复提示'),
    ('预览与清除','上传后展示缩略图/演示影像预览，支持清除后重新上传；未上传时显示暂无图像','当前原型清除按钮直接清除，不展示二次确认；生产是否二次确认待确认；不得继续引用已清除文件'),
    ('最低条件','点击开始分析时至少有一张有效影像','未满足时提示“请至少提供一种眼底影像”，不创建分析任务'),
], widths=[1.55,3.55,2.85], font_size=8.7)

# 8 analysis
doc.add_heading('8 模块需求 智能分析', level=1)
add_para(doc, '智能分析页是输入与报告之间的任务状态页。最新版原型展示四个阶段：基因与影像联合分析、关键疾病特征提取、基因与病例证据关联、综合风险评估与报告；每一阶段显示“等待分析”“分析中”“已完成”状态，并同步阶段标题、说明和进度条。当前动画由前端定时器驱动，生产版必须将视觉状态与真实任务状态绑定，不能把演示时长直接当作算法耗时或后端 SLA。')
add_table(doc, ['步骤','系统行为','可观察输出','失败处理'], [
    ('1 基因与影像联合分析','校验输入对象、建立分析上下文、记录模型与知识库版本','显示任务已接收及输入摘要','输入对象缺失或版本不可用则失败并可返回修改'),
    ('2 关键疾病特征提取','调用影像模型提取疾病相关特征；保存模型置信度和质量指标','显示影像质量/特征提取状态','低质量或模型不可用时明确标记，不伪造正常结果'),
    ('3 基因与病例证据关联','按基因状态、证据强度、适用人群和疾病方向匹配知识条目','显示证据关联进度和数据版本','知识条目缺失时标记“证据不足/未覆盖”，不静默外推'),
    ('4 综合风险评估与报告','融合影像、基因和病例证据，输出2型糖尿病、DR和UM方向的结构化风险及解释','生成报告 ID和版本，进入评估报告页','任一关键模块失败时按产品规则生成失败或部分结果报告，并显式标注缺失'),
], widths=[1.65,3.25,2.0,2.2], font_size=8.5)
add_heading('8.1 分析任务状态', level=2)
add_table(doc, ['状态','进入条件','页面行为','可转移状态'], [
    ('草稿','输入尚未提交','可编辑输入、可删除草稿','已提交/已取消/已过期'),
    ('已提交','前置校验通过并创建任务','锁定关键输入，展示任务编号','排队中/失败/取消'),
    ('排队中','等待算法服务资源','显示排队提示，可取消（若支持）','处理中/失败/取消/超时'),
    ('处理中','至少一个分析步骤执行中','按四阶段显示等待分析/分析中/已完成、阶段标题、说明和进度；禁止重复创建相同任务','生成报告/失败/超时/取消'),
    ('生成报告','全部关键步骤完成','跳转报告页，保存报告版本','已查看/已导出/重新分析'),
    ('失败/超时/取消','系统或用户终止','展示原因、任务 ID、重试/返回输入动作','重新提交/关闭'),
], widths=[1.1,2.15,3.15,1.7], font_size=8.5)
add_callout(doc, '并发规则', '同一病例同一输入上下文重复点击开始分析时，应幂等返回已有任务或明确提示已存在任务；不得因双击产生多个不可区分的分析任务。', fill='E2F0D9', color='385723')

# 9 report
doc.add_heading('9 模块需求 评估报告', level=1)
add_para(doc, '评估报告用于让医护或科研用户快速理解本次分析的输入、证据和方向性结果。报告首屏必须明确患者/病例标识、生成时间、数据来源、模型或知识库版本以及“临床辅助、不能替代诊断”的提示。报告内容按疾病方向和证据强度组织，避免将示意评分误读为发病概率。')
doc.add_heading('9.1 报告结构', level=2)
add_table(doc, ['区块','必须展示内容','可追溯信息','异常/缺失展示'], [
    ('患者背景信息','姓名/病例标识、年龄、性别、既往病史及未提供标识','输入版本、提交人、生成时间','空字段显示未提供，不用空白误导'),
    ('关键疾病基因易感性评分','按疾病方向展示遗传易感性提示、分值/等级（如上线）和解释','位点集合、证据版本、算法版本','没有足够基因输入时展示未评估或证据不足'),
    ('影像异常提示与特征说明','按影像类型和眼别展示病灶区域、特征名称、置信度/质量指标（如有）','影像对象 ID、模型版本、分析时间','影像质量不足、无法判读或无异常需区分'),
    ('基因变异解读','基因/标志物状态、证据强度、适用疾病方向、公开文献来源、效应量和人群限制；DR方向区分强/中等/较弱证据，UM方向区分驱动突变与预后相关标志物','知识条目 ID、版本、引用来源','未测、阴性、质量不合格不得混为一类；弱证据和东亚证据不足/结论不一致需明确提示'),
    ('基因变异解读小结','对本次纳入的位点或标志物进行汇总，说明哪些证据计入风险提示、哪些仅作背景参考；UM需说明驱动事件与预后指标的区别，并展示BAP1/染色体3冲突或少见共存提示','基因分型集合、知识库版本、总结生成时间','无基因输入时展示未提供/暂无解读，不得写成阴性'),
    ('影像与基因综合疾病风险评估','同时展示2型糖尿病、糖尿病视网膜病变（DR）、脉络膜黑色素瘤（UM）三类风险卡片，包括分数/等级、影像证据和基因证据','风险规则版本、输入摘要','冲突组合、证据不足或影像质量不足应输出复核提示'),
    ('方向性建议','随访/转诊/会诊方向性建议和适用前提','建议规则版本和生成时间','不得写成处方、医嘱或确定诊断'),
], widths=[1.5,3.0,2.0,1.75], font_size=8.2)
add_heading('9.2 风险分层与证据规则', level=2)
add_bullets(doc, [
    '所有风险等级必须明确是辅助提示、示意分层或经临床验证的指标；如果尚未完成临床验证，统一标注为“原型/研究阶段，不代表实际发病概率”。',
    '证据强度至少区分强证据、中等证据、较弱证据/研究阶段；弱证据不得在没有其他证据支持时单独触发高风险结论。',
    'DR方向需展示公开文献来源、效应量和人群限制；东亚人群证据不足或不同研究结论不一致时，应保留限制说明，弱证据不直接计入风险提示。',
    'UM方向需明确GNAQ/GNA11/CYSLTR2/PLCB4为驱动事件，不与BAP1、SF3B1、EIF1AX和染色体3预后信息混为一类；BAP1与染色体3不一致时提示复核。',
    '若出现BAP1阴性同时伴SF3B1或EIF1AX突变的少见共存组合，不应简单按最高风险处理；报告应建议结合GEP/PRAME等检测复核，并把该组合的处理标为待临床确认。',
    '建议文案必须由临床专家审核；原型中的3–6个月、6–12个月、年度筛查等周期不能未经确认直接固化为通用医嘱。',
])
add_para(doc, '当前原型的示意算法（仅用于复现前端演示，不得作为临床验证规则）：关键疾病基因易感性展示分数中，2型糖尿病方向约为 min(85, max(8, 18 + 强证据位点数×11 + 中等证据位点数×7 + 较弱证据位点数×3))；DR方向约为 min(80, max(8, 10 + DR中等证据位点数×14 + DR较弱位点数×5 + min(12, 2型糖尿病位点数×3))。综合风险使用的影像分数 = min(100, 20 + 有效影像数量×12 + 对应疾病方向病灶数量×8)；2型糖尿病综合风险 = 65%×2型糖尿病基因分数 + 35%×DR方向影像分数；DR综合风险 = 60%×DR基因分数 + 40%×DR方向影像分数；UM综合风险 = 60%×UM基因分数 + 40%×UM方向影像分数。原型还在DR基因分数中对年龄≥60岁增加4分。风险等级阈值为分数<30低风险、30–59中风险、≥60高风险。UM示意规则还包含BAP1阴性、SF3B1突变型、EIF1AX状态和染色体3单体/二体的分层，以及BAP1阴性与SF3B1/EIF1AX突变共存时将分数乘以0.6的降权处理。上述规则受原型输入、演示病灶和前端默认值影响，必须由算法、临床与合规团队重新确认。')
add_heading('9.3 报告导出', level=2)
add_table(doc, ['格式/入口','当前最新版原型行为','内容要求与边界','验收重点'], [
    ('PDF 综合评估报告','报告底部按钮“导出综合评估报告”；浏览器端通过独立纵向打印视图生成PDF','与页面报告一致，包含报告内容、影像缩略图、风险结果和免责声明；必须绑定报告版本','中文字体、分页、影像缩略图、敏感信息控制、失败可重试'),
    ('算法接口文档 .doc','报告底部按钮“导出算法接口文档”；浏览器端离线生成 Word 兼容的 .doc 文档','包含接口概述、请求参数、返回参数、评分规则、请求/返回示例和错误码；它是接口说明，不是患者综合评估报告','字段与规则版本、可打开性、示例与实际页面算法一致、不得误作临床报告'),
    ('Word 综合评估报告','当前最新版原型未展示该按钮或实现入口','是否保留为后续能力待确认；本版不得把它写成当前已实现功能或验收前置条件','若后续立项，另行冻结模板、签名、编号、版本和导出权限'),
], widths=[1.3,2.45,2.85,1.7], font_size=8.0)

# 10 data model
doc.add_heading('10 核心数据与接口需求', level=1)
doc.add_heading('10.1 核心对象', level=2)
add_table(doc, ['对象','关键字段','生命周期/规则'], [
    ('病例/患者上下文','case_id、患者标识、基本信息、授权状态','可被多次分析；患者直接标识与科研脱敏标识分离'),
    ('分析输入','analysis_id、case_id、基因状态集合、影像对象集合、输入版本','提交后形成不可变快照；重新分析生成新版本或新任务'),
    ('影像对象','image_id、类型、眼别、文件引用、质量状态、上传者、时间','原始文件与脱敏/缩略图分层存储；清除后不得继续被任务引用'),
    ('基因结果','gene_result_id、基因/标志物、状态、检测方法、来源、质量、证据版本','支持未提供/未检测/阴性/阳性/异常/不确定等状态'),
    ('分析任务','task_id、状态、步骤、错误码、模型版本、开始/结束时间','状态可查询、可重试、可取消；保留审计轨迹'),
    ('评估报告','report_id、analysis_id、报告版本、风险结果、解释、建议、免责声明','生成后内容不可静默覆盖；修订需形成新版本'),
], widths=[1.25,3.55,3.5], font_size=8.8)
doc.add_heading('10.2 服务接口业务契约', level=2)
add_para(doc, '原型意图支持算法接口调用，但未提供已确认的接口地址、认证方式和最终字段。生产接口至少需要以下业务输入输出；字段名、数据类型和协议由技术方案另行冻结。')
add_table(doc, ['接口','请求业务输入','成功输出','失败输出'], [
    ('创建分析任务','病例/患者授权上下文、输入对象引用、基因状态、影像引用、模型/知识库版本（可选）','task_id、analysis_id、已接受时间、初始状态','参数错误、无权限、重复任务、对象不存在、服务不可用'),
    ('查询分析任务','task_id或analysis_id、访问者身份','状态、步骤进度、当前提示、错误码、report_id','任务不存在、无权限、状态不可查询'),
    ('获取报告','report_id、访问者身份、期望版本','结构化报告、解释、免责声明、版本信息','报告未生成、已删除、无权限、版本不存在'),
    ('导出综合评估报告（PDF）','report_id、报告版本、访问者身份','PDF文件引用、导出状态、版本和审计标识','无权限、内容超限、生成失败、中文字体或分页失败'),
    ('导出算法接口文档（.doc）','接口文档版本、字段范围、访问者身份','.doc文件引用、文档版本、导出状态','字段未冻结、生成失败、无权限；不得与患者报告导出混用'),
], widths=[1.35,3.25,2.4,2.1], font_size=8.5)

# 11 exceptions/security
doc.add_heading('11 异常、边界与安全需求', level=1)
add_table(doc, ['场景','系统行为','用户可见反馈','验收要求'], [
    ('无影像点击分析','不创建任务，焦点定位上传区','提示至少提供一种眼底影像','不会进入分析页，不产生空报告'),
    ('影像上传失败','不保存无效对象；保留其他有效输入','说明失败原因并支持重试','已清除/失败文件不能出现在报告中'),
    ('影像质量不足','按模型能力返回不可判读或低质量','显示质量原因和补拍/重传建议','不自动输出正常或低风险结论'),
    ('基因结果为空','按产品确认规则进入影像分析或阻断','明确“未提供基因结果”','报告不得把缺失当阴性'),
    ('基因状态冲突','阻止提交或进入复核态','指出互斥项和修正位置','同一标志物不能同时呈现相反状态'),
    ('算法超时/失败','任务进入失败/超时，记录错误码','展示任务 ID、重试或返回输入动作','不显示旧任务结果冒充新结果'),
    ('模型/知识库版本失效','拒绝或降级到被允许版本','提示版本不可用或结果需复核','报告记录实际使用版本'),
    ('权限失效/会话过期','停止读取敏感数据，要求重新认证','统一提示，不暴露患者信息','接口层也必须拦截'),
    ('重复提交/并发','按幂等键合并或拒绝重复任务','提示已有任务正在处理','最多形成一个可识别的任务结果'),
    ('PDF导出失败','保留报告本身，PDF导出任务可重试','显示失败原因，不生成损坏文件；不影响报告查看','失败文件不可下载，必须记录导出失败审计'),
    ('算法接口文档导出失败','保留报告本身，接口文档可重新生成','提示文档生成失败，不影响临床报告查看；不得以空文件代替','失败文件不可下载，文档版本与报告/算法规则版本可追溯'),
], widths=[1.45,2.65,2.3,2.0], font_size=8.2)
doc.add_heading('11.1 数据安全与审计', level=2)
add_bullets(doc, [
    '患者直接身份信息、基因结果和原始影像属于敏感数据，访问、下载、导出和删除必须受权限控制并记录审计。',
    '科研场景默认使用脱敏病例标识；导出前必须校验字段范围，不能通过报告模板绕过脱敏。',
    '审计日志至少记录操作者、组织/角色、对象 ID、操作类型、时间、结果、失败原因、模型/知识库版本和导出文件标识。',
    '影像上传需执行文件类型校验、病毒扫描、元数据处理和安全存储；具体保留期限、备份和删除策略待确认。',
])

# 12 nonfunctional
doc.add_heading('12 非功能需求', level=1)
add_table(doc, ['类别','需求'], [
    ('可用性','输入、分析和报告页面在桌面及常见平板宽度下可用；错误提示靠近问题位置；关键操作有明确成功/失败反馈'),
    ('可靠性','分析任务状态可查询、可重试、可恢复；浏览器刷新不导致任务状态丢失；重复提交具备幂等策略'),
    ('可解释性','报告展示输入摘要、证据强度、限制条件、模型/知识库版本和免责声明；不能只输出单一分数'),
    ('可追溯性','报告与输入、影像、基因状态、规则版本和操作者建立关联；重新分析不覆盖历史报告'),
    ('性能','具体时延目标待确认；系统至少应在前端显示排队/处理中状态，并避免长时间无反馈'),
    ('兼容性','支持主流现代浏览器；上传文件、导出文件和中文字体在目标环境中可正常处理'),
    ('可维护性','疾病、基因、证据、提示语和风险规则尽量配置化并具备版本管理，避免写死在页面脚本中'),
], widths=[1.15,7.25], font_size=8.8)

# 13 acceptance
doc.add_heading('13 验收标准', level=1)
add_para(doc, '验收以“前置条件—操作—预期结果”为准。以下用例覆盖当前原型链路的正常、异常、权限和状态路径；正式上线前应由临床、算法和安全团队补充真实数据集与模型验证标准。')
add_table(doc, ['编号','前置条件','操作','预期结果'], [
    ('AC-01','进入输入页，未上传影像','直接点击开始一键分析','页面提示至少提供一种眼底影像；不创建任务'),
    ('AC-02','配置眼底彩照+单眼','上传一张有效图片并发起分析','生成任务，进入智能分析页，影像类型和眼别传递正确'),
    ('AC-03','配置OCT+OCTA双眼（按产品允许）','分别上传右眼/左眼文件','每个槽位独立预览、清除、校验和传递；报告按眼别展示'),
    ('AC-04','选择同一标志物的相反状态','先勾选阳性，再勾选阴性/野生型','系统自动互斥或阻止提交，并保留明确提示'),
    ('AC-05','未输入基因结果但有有效影像','发起分析','按已确认规则进入影像分析或提示需补充；报告明确基因未提供，不显示为阴性'),
    ('AC-06','分析处理中刷新页面','刷新并重新进入任务','系统根据任务 ID恢复真实状态，不从头伪造进度'),
    ('AC-07','算法服务返回低质量影像','查看报告','报告标记不可判读/低质量，不能自动输出正常或低风险结论'),
    ('AC-08','基因和影像证据方向冲突','查看综合风险','报告展示冲突与复核提示，不强行输出确定性诊断'),
    ('AC-09','分析成功','查看报告全部区块','患者背景、基因易感性、影像提示、基因解读、综合风险和建议均有内容或明确缺失态'),
    ('AC-10','用户无报告访问权限','通过页面和接口获取报告','页面和接口均拒绝访问，不泄露患者姓名、影像或基因信息'),
    ('AC-11','报告成功生成','点击“导出综合评估报告”','生成与当前报告版本一致的PDF，包含报告内容、影像缩略图、风险结果和免责声明；失败时不产生可下载损坏文件'),
    ('AC-12','报告成功生成','点击“导出算法接口文档”','生成可打开的 .doc 接口说明文档，包含请求/返回字段、示例、评分规则和错误码；明确其不是患者综合评估报告'),
    ('AC-13','用户重复点击开始分析','短时间内连续点击两次','只创建一个可识别任务，页面给出处理中/已提交反馈'),
    ('AC-14','点击重新分析','从报告返回输入页','不得静默覆盖原报告；新任务或新版本可追溯到原任务'),
], widths=[0.7,2.1,2.8,3.0], font_size=8.1)

# 14 assumptions and change log
doc.add_heading('14 版本变更与待产品确认汇总', level=1)
doc.add_heading('14.1 本次版本变更记录', level=2)
add_table(doc, ['编号','变更项','最新版原型事实','对需求文档的影响'], [
    ('CHG-01','基因录入方式','统一多选下拉面板；新增“已选择的基因分型”摘要；BAP1/SF3B1/EIF1AX/染色体3状态互斥','补充控件、互斥、摘要和缺失语义边界'),
    ('CHG-02','影像输入','支持CFP、OCT、OCTA类型选择；支持单眼/双眼；右眼/左眼独立槽位；可上传/清除本地图像','补充影像类型、眼别、槽位和清除规则'),
    ('CHG-03','智能分析','四阶段流水线，每步显示等待分析/分析中/已完成，并显示阶段说明和进度条','补充阶段状态；明确动画不等于生产算法状态'),
    ('CHG-04','报告内容','新增基因变异解读小结；风险区同时展示2型糖尿病、DR、UM；影像与文字采用左右布局','补充报告结构、证据分层和多疾病风险卡片'),
    ('CHG-05','UM风险解读','新增驱动突变与预后标志物区分、BAP1/染色体3冲突提示、BAP1阴性与SF3B1/EIF1AX突变共存提示及GEP/PRAME复核建议','补充UM规则、冲突和待临床确认边界'),
    ('CHG-06','导出能力','当前提供PDF综合评估报告和算法接口文档 .doc；未展示Word综合评估报告入口','删除Word综合报告已实现的表述，拆分两种当前导出能力'),
], widths=[0.75,1.45,3.35,2.15], font_size=8.0)
doc.add_heading('14.2 工作假设与待产品确认汇总', level=2)
add_table(doc, ['编号','事项','本文暂采用的工作假设','需要谁确认'], [
    ('Q-01','服务对象','以医护/科研用户为主，不直接面向患者作最终诊断展示','产品、临床、合规'),
    ('Q-02','基因状态','生产数据模型区分未提供、未检测、阴性、阳性、异常和不确定；原型“未勾选=未检测/阴性”仅视为原型现状','产品、遗传/检验、算法'),
    ('Q-03','疾病范围','本期只覆盖原型中的糖尿病/糖尿病视网膜病变和脉络膜黑色素瘤方向','产品、临床'),
    ('Q-04','评分性质','当前分值与阈值为原型示意，不作为临床验证风险工具','临床、算法、合规'),
    ('Q-05','真实算法','生产版调用独立算法服务，并返回可追溯的模型版本、质量指标和错误状态','算法、后端'),
    ('Q-06','报告导出','当前原型明确展示PDF综合评估报告和算法接口文档 .doc 两种导出；PDF模板、签名、编号、算法文档字段及版本规则仍待确定','产品、研发、算法、合规'),
    ('Q-07','Word综合评估报告','当前最新版原型未展示该能力，本文不将其作为当前实现；如需保留，应纳入后续版本评审','产品、研发、合规'),
    ('Q-08','数据治理','敏感数据采用最小权限、审计、脱敏和生命周期管理','安全、运维、合规'),
], widths=[0.7,1.65,4.55,1.7], font_size=8.3)

# 15 checklist
doc.add_heading('15 交付前完整性检查', level=1)
add_bullets(doc, [
    '范围已限定为基于基因的眼病/慢病诊断应用，未混入其他产品的需求。',
    '已覆盖输入、分析、报告三个页面状态，以及跨页面数据传递、刷新、重试、取消和重新分析。',
    '已覆盖患者信息、基因/标志物、CFP/OCT/OCTA眼底影像、任务、报告、PDF综合报告和算法接口文档 .doc 等核心对象。',
    '已区分原型事实、工作假设和待产品确认项；未将演示数据、示意分数和定时动画直接固化为临床规则。',
    '已覆盖角色权限、隐私安全、异常恢复、幂等、版本追溯和报告免责声明。',
    '已给出可观察的验收用例，包含正常、异常、权限、分析阶段和两类导出路径。',
])
add_callout(doc, '上线前阻塞项', '在正式上线前，必须完成疾病与基因知识库审核、影像模型临床验证、风险分层规则审批、数据安全与隐私评审，以及PDF综合评估报告模板和算法接口文档版本规则确认。Word版综合评估报告仅在产品确认保留后另行立项。', fill='F4CCCC', color='990000')

# Save
# Ensure core properties
props = doc.core_properties
props.title = '基于基因的眼病慢病诊断应用需求规格说明'
props.subject = '基于基因、眼底影像和临床背景的诊断辅助应用需求'
props.author = 'Codex'
props.comments = '基于最新版 HTML 原型（2026-09-20）对原需求文档更新；含原型事实、工作假设和待确认项'
doc.save(OUT)
print(str(OUT))
