# -*- coding: utf-8 -*-
"""
情境默写抽题工具 —— 从 Excel 题库中抽取题目，生成 Word 或 PPT 文档。
支持多篇目勾选、自定义抽题数量、顺序/乱序排列。
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import random
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Cm as PptCm, Pt as PptPt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


class App:
    """主应用程序类，管理 GUI 和文档生成逻辑"""

    def __init__(self, root):
        self.root = root
        self.root.title("情境默写抽题工具")
        self.root.geometry("800x600")

        # 数据存储
        self.df = None               # 整个 Excel 数据框
        self.sheet_names = []        # 所有篇目名称列表
        self.topic_counts = {}       # 每个篇目的总题数
        self.selected_topics = {}    # 勾选状态字典 {篇目名: BooleanVar}
        self.input_vars = {}         # 输入数量字典 {篇目名: StringVar}

        # 模式选择
        self.mode = tk.StringVar(value="docx")       # 输出格式：docx / pptx
        self.order_mode = tk.StringVar(value="顺序")  # 排序方式：顺序 / 乱序

        self.create_widgets()

    # ------------------------------ GUI 构建 ------------------------------
    def create_widgets(self):
        """创建所有 GUI 控件"""
        # 顶部：文件选择
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top_frame, text="选择Excel文件:").pack(side=tk.LEFT)
        self.file_path_label = ttk.Label(top_frame, text="未选择", foreground="gray")
        self.file_path_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="浏览...", command=self.select_file).pack(side=tk.LEFT)

        # 输出模式选择
        mode_frame = ttk.Frame(self.root)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(mode_frame, text="输出模式:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Word (.docx)", variable=self.mode, value="docx").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="PPT (.pptx)", variable=self.mode, value="pptx").pack(side=tk.LEFT, padx=5)

        # 中间滚动区域：显示篇目列表
        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas = tk.Canvas(middle_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(middle_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 底部：排序方式和操作按钮
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)

        order_frame = ttk.Frame(bottom_frame)
        order_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(order_frame, text="排序方式:").pack(side=tk.LEFT)
        ttk.Radiobutton(order_frame, text="顺序", variable=self.order_mode, value="顺序").pack(side=tk.LEFT)
        ttk.Radiobutton(order_frame, text="乱序", variable=self.order_mode, value="乱序").pack(side=tk.LEFT)

        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="生成文档", command=self.generate_output).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT)

        # 初始禁用所有篇目控件（加载文件后才启用）
        self.set_controls_state(tk.DISABLED)

    def set_controls_state(self, state):
        """统一设置滚动区域内所有 Checkbutton 和 Entry 的状态"""
        for child in self.scrollable_frame.winfo_children():
            if isinstance(child, ttk.Checkbutton):
                child.config(state=state)
            elif isinstance(child, ttk.Entry):
                child.config(state=state)

    # ------------------------------ 文件加载 ------------------------------
    def select_file(self):
        """弹出文件选择对话框，加载 Excel 题库"""
        file_path = filedialog.askopenfilename(
            title="选择Excel题库文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            # 读取 Excel，第一行为表头
            self.df = pd.read_excel(file_path, engine='openpyxl', header=0)
            if self.df.shape[1] < 5:
                raise ValueError("Excel必须包含至少5列（篇目、题干、答案A、答案B、答案C）")

            # 提取篇目名称（第一列去重）
            self.sheet_names = self.df.iloc[:, 0].dropna().unique().tolist()
            self.sheet_names = [str(name).strip() for name in self.sheet_names if str(name).strip() != '']

            # 统计每个篇目的题数
            self.topic_counts = {}
            for name in self.sheet_names:
                count = len(self.df[self.df.iloc[:, 0] == name])
                self.topic_counts[name] = count

            # 刷新 UI 列表
            self.refresh_topic_list()
            self.file_path_label.config(text=os.path.basename(file_path), foreground="black")
            self.set_controls_state(tk.NORMAL)   # 启用控件
        except Exception as e:
            messagebox.showerror("文件读取失败", str(e))

    def refresh_topic_list(self):
        """根据 sheet_names 重新绘制篇目选择列表"""
        # 清空旧控件
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.selected_topics.clear()
        self.input_vars.clear()

        # 表头行
        header_frame = ttk.Frame(self.scrollable_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(header_frame, text="勾选", width=5).grid(row=0, column=0)
        ttk.Label(header_frame, text="篇目", width=30, anchor=tk.W).grid(row=0, column=1)
        ttk.Label(header_frame, text="现有题量", width=8).grid(row=0, column=2)
        ttk.Label(header_frame, text="选择数量", width=10).grid(row=0, column=3)

        # 每行一个篇目
        for idx, name in enumerate(self.sheet_names):
            row_frame = ttk.Frame(self.scrollable_frame)
            row_frame.pack(fill=tk.X, padx=5, pady=1)

            var_check = tk.BooleanVar()
            chk = ttk.Checkbutton(row_frame, variable=var_check)
            chk.grid(row=0, column=0)
            self.selected_topics[name] = var_check

            lbl_name = ttk.Label(row_frame, text=name, width=30, anchor=tk.W)
            lbl_name.grid(row=0, column=1, padx=(5,0))

            count = self.topic_counts[name]
            lbl_count = ttk.Label(row_frame, text=str(count), width=8)
            lbl_count.grid(row=0, column=2)

            var_input = tk.StringVar(value="0")
            entry = ttk.Entry(row_frame, textvariable=var_input, width=10, justify=tk.CENTER)
            entry.grid(row=0, column=3, padx=(5,0))
            self.input_vars[name] = var_input

    # ------------------------------ 验证与抽题 ------------------------------
    def validate_selection(self):
        """
        检查用户勾选和输入是否合法，并从 DataFrame 中随机抽取题目。
        返回: (是否合法, 错误信息字符串, 抽取的题目列表)
        """
        selected_items = []
        errors = []
        has_checked = False

        for name, var in self.selected_topics.items():
            if var.get():   # 该篇目被勾选
                has_checked = True
                input_str = self.input_vars[name].get().strip()
                # 输入必须为正整数
                if not input_str.isdigit() or int(input_str) <= 0:
                    errors.append(f"篇目「{name}」的选择数量无效，请输入正整数。")
                    continue
                num = int(input_str)
                max_num = self.topic_counts[name]
                if num > max_num:
                    errors.append(f"篇目「{name}」选择数量 {num} 超过现有题量 {max_num}。")
                    continue

                topic_df = self.df[self.df.iloc[:, 0] == name]
                if num == max_num:
                    chosen = topic_df               # 全部选取
                else:
                    chosen = topic_df.sample(n=num, random_state=None)   # 随机抽取

                # 将选中行转换为字典列表
                for _, row in chosen.iterrows():
                    selected_items.append({
                        'topic': name,
                        'stem': str(row.iloc[1]),                         # 题干（含A/B/C占位符）
                        'answer_a': str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
                        'answer_b': str(row.iloc[3]) if pd.notna(row.iloc[3]) else "",
                        'answer_c': str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
                    })

        if not has_checked:
            errors.append("请至少勾选一个篇目。")

        if errors:
            return False, "\n".join(errors), []
        return True, "", selected_items

    # ------------------------------ Word 换行辅助 ------------------------------
    def force_wordwrap_on_paragraph(self, paragraph):
        """强制段落自动换行（解决某些情况下不换行的问题）"""
        pPr = paragraph._element.get_or_add_pPr()
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        # 移除可能导致不换行的属性
        for tag in ['w:noWrap', 'w:keepLines']:
            for el in pPr.findall(tag, nsmap):
                pPr.remove(el)
        # 添加 wordWrap 属性
        ww = pPr.find('w:wordWrap', nsmap)
        if ww is None:
            ww = OxmlElement('w:wordWrap')
            pPr.insert(0, ww)
        ww.set(qn('w:val'), '1')
        # 关闭自动间距
        for tag, val in [('w:autoSpaceDE', '0'), ('w:autoSpaceDN', '0')]:
            el = pPr.find(tag, nsmap)
            if el is None:
                el = OxmlElement(tag)
                pPr.append(el)
            el.set(qn('w:val'), val)

    def set_paragraph_format(self, paragraph):
        """设置段落行距并强制换行"""
        paragraph.paragraph_format.line_spacing = 1.5
        self.force_wordwrap_on_paragraph(paragraph)

    def apply_global_wordwrap(self, doc):
        """对 Normal 样式全局应用自动换行"""
        styles_element = doc.styles.element
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        normal_style = styles_element.find('.//w:style[@w:type="paragraph"][@w:styleId="Normal"]', nsmap)
        if normal_style is not None:
            pPr = normal_style.find('w:pPr', nsmap)
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                normal_style.insert(0, pPr)
            ww = pPr.find('w:wordWrap', nsmap)
            if ww is None:
                ww = OxmlElement('w:wordWrap')
                pPr.insert(0, ww)
            ww.set(qn('w:val'), '1')
            for tag in ['w:noWrap', 'w:keepLines']:
                for el in pPr.findall(tag, nsmap):
                    pPr.remove(el)

    # ------------------------------ 核心：构建题目与答案 ------------------------------
    def _build_question_and_answer(self, items, underline_len=25):
        """
        根据抽取的题目列表，生成用于显示的题目字符串和带分段标记的答案列表。
        :param items: 题目字典列表
        :param underline_len: 下划线长度，docx默认25，pptx传4（可根据需要调整）
        :return: (questions, answers_parts)
                 questions: 每道题的完整字符串（含编号、下划线）
                 answers_parts: 答案的分段列表，每段为 (文本, 是否带下划线) 的元组列表
        """
        questions = []
        answers_parts = []

        for idx, item in enumerate(items, start=1):
            stem = item['stem']
            ans_a = item['answer_a']
            ans_b = item['answer_b']
            ans_c_raw = item['answer_c']
            ans_c = ans_c_raw.strip() if ans_c_raw and ans_c_raw.strip() else ''

            # ----- 生成题目（删除C占位符，A/B替换为指定长度的下划线）-----
            q_chars = []
            for ch in stem:
                if ch == 'A':
                    q_chars.append('_' * underline_len)      # 使用变量长度
                elif ch == 'B':
                    q_chars.append('_' * underline_len)
                elif ch == 'C':
                    continue          # 直接删除C占位符
                else:
                    q_chars.append(ch)
            question_str = ''.join(q_chars)
            questions.append(f"{idx}. {question_str}")

            # ----- 生成答案分段（下划线部分保留原文答案）-----
            placeholder_map = {
                'A': ans_a,
                'B': ans_b,
                'C': f"（{ans_c}）" if ans_c else ''   # 空则删除
            }
            # 找出所有占位符位置
            positions = [(i, ch) for i, ch in enumerate(stem) if ch in placeholder_map]
            parts = []
            last_end = 0
            for pos, ch in positions:
                if pos > last_end:
                    parts.append((stem[last_end:pos], False))   # 普通文本
                answer_text = placeholder_map[ch]
                if answer_text:                                 # 只有非空才添加带下划线的片段
                    parts.append((answer_text, True))           # 下划线标记
                last_end = pos + 1
            if last_end < len(stem):
                parts.append((stem[last_end:], False))
            # 插入题号前缀
            parts.insert(0, (f"{idx}. ", False))
            answers_parts.append(parts)

        return questions, answers_parts

    # ------------------------------ 输出入口 ------------------------------
    def generate_output(self):
        """根据当前选择的输出模式调用对应的生成函数"""
        if self.mode.get() == "docx":
            self.generate_docx()
        elif self.mode.get() == "pptx":
            self.generate_pptx()

    # ------------------------------ 生成 Word 文档 ------------------------------
    def generate_docx(self):
        """生成 .docx 文件，包含题目页和答案页"""
        valid, err_msg, items = self.validate_selection()
        if not valid:
            messagebox.showerror("输入错误", err_msg)
            return
        if not items:
            messagebox.showwarning("无题目", "没有符合条件的题目可生成。")
            return

        # 排序：顺序模式下按篇目原顺序排列，乱序则打乱
        if self.order_mode.get() == "顺序":
            ordered = []
            for name in self.sheet_names:
                group = [it for it in items if it['topic'] == name]
                ordered.extend(group)
            items = ordered
        else:
            random.shuffle(items)

        # 构建题目和答案（docx 使用默认下划线长度25）
        questions, answers_parts = self._build_question_and_answer(items)

        # 创建 Word 文档对象
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(1.27)
        section.right_margin = Cm(1.27)

        # 设置默认样式
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(10.5)
        rFonts = style.element.rPr.rFonts if style.element.rPr is not None else None
        if rFonts is None:
            rPr = style.element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:eastAsia'), '宋体')
        style.paragraph_format.line_spacing = 1.5
        self.apply_global_wordwrap(doc)

        # ---- 题目页 ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("情境默写")
        r.bold = True
        r.font.size = Pt(16)
        r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.set_paragraph_format(p)

        for q in questions:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(q)
            r.font.size = Pt(10.5)
            r.font.name = 'Times New Roman'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            self.set_paragraph_format(p)

        # 分页
        doc.add_page_break()

        # ---- 答案页 ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("答案")
        r.bold = True
        r.font.size = Pt(16)
        r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.set_paragraph_format(p)

        for parts in answers_parts:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            self.set_paragraph_format(p)
            for text, underline in parts:
                r = p.add_run(text)
                r.font.size = Pt(10.5)
                r.font.name = 'Times New Roman'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                if underline:
                    r.underline = True

        # 保存文档
        save_dir = "C:\\"
        filename = f"情境默写_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx"
        full_path = os.path.join(save_dir, filename)
        try:
            doc.save(full_path)
            messagebox.showinfo("成功", f"文档已保存至：\n{full_path}")
        except PermissionError:
            new_path = filedialog.asksaveasfilename(
                title="保存文档",
                defaultextension=".docx",
                initialfile=filename,
                filetypes=[("Word文档", "*.docx")]
            )
            if new_path:
                try:
                    doc.save(new_path)
                    messagebox.showinfo("成功", f"文档已保存至：\n{new_path}")
                except Exception as e2:
                    messagebox.showerror("保存失败", f"无法保存文件：{str(e2)}")
            else:
                messagebox.showwarning("取消", "用户取消了保存操作。")
        except Exception as e:
            messagebox.showerror("保存失败", f"发生未知错误：{str(e)}")

    # ------------------------------ 生成 PPT 文档 ------------------------------
    def generate_pptx(self):
        """生成 .pptx 文件，每页最多5道题/答案"""
        valid, err_msg, items = self.validate_selection()
        if not valid:
            messagebox.showerror("输入错误", err_msg)
            return
        if not items:
            messagebox.showwarning("无题目", "没有符合条件的题目可生成。")
            return

        # 排序
        if self.order_mode.get() == "顺序":
            ordered = []
            for name in self.sheet_names:
                group = [it for it in items if it['topic'] == name]
                ordered.extend(group)
            items = ordered
        else:
            random.shuffle(items)

        # 构建题目和答案（pptx 使用较短的下划线，长度为4）
        questions, answers_parts = self._build_question_and_answer(items, underline_len=4)

        # 创建 PPT 演示文稿
        prs = Presentation()
        prs.slide_width = PptCm(33)
        prs.slide_height = PptCm(16)
        blank_layout = prs.slide_layouts[6]   # 空白版式
        if blank_layout is None:
            blank_layout = prs.slide_layouts[0]

        def add_textbox_slide(content_list, is_answer=False):
            """
            添加一张幻灯片，内容由 content_list 决定。
            :param content_list: 如果是题目页，则为字符串列表；如果是答案页，则为分段列表。
            :param is_answer: 是否为答案页
            """
            slide = prs.slides.add_slide(blank_layout)
            # 移除原有占位符
            for ph in slide.placeholders:
                sp = ph._element
                sp.getparent().remove(sp)

            # 【修改】文本框尺寸：31.5cm宽 × 13.5cm高，上边界距页面顶部2cm
            txBox = slide.shapes.add_textbox(
                left=PptCm(0),       # 左边界0cm
                top=PptCm(2),        # 上边界2cm（可调整）
                width=PptCm(31.5),   # 宽度31.5cm（可调整）
                height=PptCm(13.5)   # 高度13.5cm（可调整）
            )
            tf = txBox.text_frame
            tf.word_wrap = True

            if is_answer:
                # 答案页：每项是一个分段列表，逐段添加 run
                for parts in content_list:
                    p = tf.add_paragraph()
                    p.alignment = PP_ALIGN.LEFT
                    for text, is_ans in parts:
                        run = p.add_run()
                        run.text = text
                        run.font.size = PptPt(24)
                        try:
                            run.font.name = '华文中宋'
                        except:
                            run.font.name = '宋体'
                        if is_ans:
                            run.font.underline = True
                            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)   # 红色下划线答案
                        else:
                            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)   # 黑色正文
                    p.space_after = PptPt(6)
            else:
                # 题目页：每项是一个纯文本字符串
                for line in content_list:
                    p = tf.add_paragraph()
                    p.alignment = PP_ALIGN.LEFT
                    run = p.add_run()
                    run.text = line
                    run.font.size = PptPt(24)
                    try:
                        run.font.name = '华文中宋'
                    except:
                        run.font.name = '宋体'
                    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                    p.space_after = PptPt(6)

            # 移除第一个空段落（如果有的话）
            if tf.paragraphs[0].text == '':
                p_elem = tf.paragraphs[0]._pPr
                if p_elem is not None:
                    p_elem.getparent().remove(p_elem)

        # 分页：每页最多5道题
        for i in range(0, len(questions), 5):
            chunk = questions[i:i+5]
            add_textbox_slide(chunk, is_answer=False)

        for i in range(0, len(answers_parts), 5):
            chunk = answers_parts[i:i+5]
            add_textbox_slide(chunk, is_answer=True)

        # 保存 PPT
        save_dir = "C:\\"
        filename = f"情境默写_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        full_path = os.path.join(save_dir, filename)
        try:
            prs.save(full_path)
            messagebox.showinfo("成功", f"PPT已保存至：\n{full_path}")
        except PermissionError:
            new_path = filedialog.asksaveasfilename(
                title="保存PPT",
                defaultextension=".pptx",
                initialfile=filename,
                filetypes=[("PowerPoint文件", "*.pptx")]
            )
            if new_path:
                try:
                    prs.save(new_path)
                    messagebox.showinfo("成功", f"PPT已保存至：\n{new_path}")
                except Exception as e2:
                    messagebox.showerror("保存失败", f"无法保存文件：{str(e2)}")
            else:
                messagebox.showwarning("取消", "用户取消了保存操作。")
        except Exception as e:
            messagebox.showerror("保存失败", f"发生未知错误：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()