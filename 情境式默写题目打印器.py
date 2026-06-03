import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import random
import os
import subprocess
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree  # 需要安装 lxml：pip install lxml
from pptx import Presentation
from pptx.util import Cm as PptCm, Pt as PptPt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("情境默写抽题工具")
        self.root.geometry("800x600")

        self.df = None
        self.sheet_names = []
        self.topic_counts = {}
        self.selected_topics = {}
        self.input_vars = {}

        self.mode = tk.StringVar(value="docx")
        self.order_mode = tk.StringVar(value="顺序")

        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top_frame, text="选择Excel文件:").pack(side=tk.LEFT)
        self.file_path_label = ttk.Label(top_frame, text="未选择", foreground="gray")
        self.file_path_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="浏览...", command=self.select_file).pack(side=tk.LEFT)

        mode_frame = ttk.Frame(self.root)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(mode_frame, text="输出模式:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Word (.docx)", variable=self.mode, value="docx").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="PPT (.pptx)", variable=self.mode, value="pptx").pack(side=tk.LEFT, padx=5)

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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

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

        self.set_controls_state(tk.DISABLED)

    def set_controls_state(self, state):
        for child in self.scrollable_frame.winfo_children():
            if isinstance(child, ttk.Checkbutton):
                child.config(state=state)
            elif isinstance(child, ttk.Entry):
                child.config(state=state)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel题库文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            self.df = pd.read_excel(file_path, engine='openpyxl', header=0)
            if self.df.shape[1] < 4:
                raise ValueError("Excel必须包含至少4列（篇目、题干、答案A、答案B）")
            self.sheet_names = self.df.iloc[:, 0].dropna().unique().tolist()
            self.sheet_names = [str(name).strip() for name in self.sheet_names if str(name).strip() != '']
            self.topic_counts = {}
            for name in self.sheet_names:
                count = len(self.df[self.df.iloc[:, 0] == name])
                self.topic_counts[name] = count
            self.refresh_topic_list()
            self.file_path_label.config(text=os.path.basename(file_path), foreground="black")
            self.set_controls_state(tk.NORMAL)
        except Exception as e:
            messagebox.showerror("文件读取失败", str(e))

    def refresh_topic_list(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.selected_topics.clear()
        self.input_vars.clear()

        header_frame = ttk.Frame(self.scrollable_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(header_frame, text="勾选", width=5).grid(row=0, column=0)
        ttk.Label(header_frame, text="篇目", width=30, anchor=tk.W).grid(row=0, column=1)
        ttk.Label(header_frame, text="现有题量", width=8).grid(row=0, column=2)
        ttk.Label(header_frame, text="选择数量", width=10).grid(row=0, column=3)

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

    def validate_selection(self):
        selected_items = []
        errors = []
        has_checked = False
        for name, var in self.selected_topics.items():
            if var.get():
                has_checked = True
                input_str = self.input_vars[name].get().strip()
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
                    chosen = topic_df
                else:
                    chosen = topic_df.sample(n=num, random_state=None)
                for _, row in chosen.iterrows():
                    selected_items.append({
                        'topic': name,
                        'stem': str(row.iloc[1]),
                        'answer_a': str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
                        'answer_b': str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
                    })
        if not has_checked:
            errors.append("请至少勾选一个篇目。")
        if errors:
            return False, "\n".join(errors), []
        return True, "", selected_items

    def force_wordwrap_on_paragraph(self, paragraph):
        """强制段落允许西文在单词中间换行（直接操作XML）"""
        pPr = paragraph._element.get_or_add_pPr()
        # 命名空间映射
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # 删除所有 noWrap 和 keepLines
        for tag in ['w:noWrap', 'w:keepLines']:
            for el in pPr.findall(tag, nsmap):
                pPr.remove(el)

        # 设置 wordWrap = 1
        ww = pPr.find('w:wordWrap', nsmap)
        if ww is None:
            ww = OxmlElement('w:wordWrap')
            pPr.insert(0, ww)  # 插入到最前面
        ww.set(qn('w:val'), '1')

        # 关闭自动调整中文与西文/数字间距（可选）
        for tag, val in [('w:autoSpaceDE', '0'), ('w:autoSpaceDN', '0')]:
            el = pPr.find(tag, nsmap)
            if el is None:
                el = OxmlElement(tag)
                pPr.append(el)
            el.set(qn('w:val'), val)

    def set_paragraph_format(self, paragraph):
        """设置1.5倍行距并强制允许西文换行"""
        paragraph.paragraph_format.line_spacing = 1.5
        self.force_wordwrap_on_paragraph(paragraph)

    def apply_global_wordwrap(self, doc):
        """修改文档的Normal样式，使其默认允许西文换行"""
        styles_element = doc.styles.element
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        # 找到 Normal 样式
        normal_style = styles_element.find('.//w:style[@w:type="paragraph"][@w:styleId="Normal"]', nsmap)
        if normal_style is not None:
            pPr = normal_style.find('w:pPr', nsmap)
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                normal_style.insert(0, pPr)
            # 同样设置 wordWrap
            ww = pPr.find('w:wordWrap', nsmap)
            if ww is None:
                ww = OxmlElement('w:wordWrap')
                pPr.insert(0, ww)
            ww.set(qn('w:val'), '1')
            # 删除 noWrap
            for tag in ['w:noWrap', 'w:keepLines']:
                for el in pPr.findall(tag, nsmap):
                    pPr.remove(el)

    def generate_output(self):
        if self.mode.get() == "docx":
            self.generate_docx()
        elif self.mode.get() == "pptx":
            self.generate_pptx()

    def generate_docx(self):
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

        # 构建题目和答案
        questions = []
        answers_parts = []
        for idx, item in enumerate(items, start=1):
            stem = item['stem']
            question_stem = stem.replace('A', '_________________________').replace('B', '_________________________')
            questions.append(f"{idx}. {question_stem}")

            pos_a = stem.find('A')
            pos_b = stem.find('B')
            if pos_a > pos_b:
                pos_a, pos_b = pos_b, pos_a
                item['answer_a'], item['answer_b'] = item['answer_b'], item['answer_a']
            prefix = stem[:pos_a]
            middle = stem[pos_a+1:pos_b]
            suffix = stem[pos_b+1:]
            parts = [
                (f"{idx}. ", False),
                (prefix, False),
                (item['answer_a'], True),
                (middle, False),
                (item['answer_b'], True),
                (suffix, False)
            ]
            answers_parts.append(parts)

        # 创建文档
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(1.27)
        section.right_margin = Cm(1.27)

        # 全局样式设置
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

        # 应用全局 wordWrap 到 Normal 样式
        self.apply_global_wordwrap(doc)

        # 标题：情境默写
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

        doc.add_page_break()

        # 答案标题
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

        # 保存
        save_dir = "C:\\"
        filename = f"情境默写_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx"
        full_path = os.path.join(save_dir, filename)
        try:
            doc.save(full_path)
            messagebox.showinfo("成功", f"文档已保存至：\n{full_path}")
            subprocess.Popen(['explorer', '/select,', full_path])
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
                    subprocess.Popen(['explorer', '/select,', new_path])
                except Exception as e2:
                    messagebox.showerror("保存失败", f"无法保存文件：{str(e2)}")
            else:
                messagebox.showwarning("取消", "用户取消了保存操作。")
        except Exception as e:
            messagebox.showerror("保存失败", f"发生未知错误：{str(e)}")

    def generate_pptx(self):
        valid, err_msg, items = self.validate_selection()
        if not valid:
            messagebox.showerror("输入错误", err_msg)
            return
        if not items:
            messagebox.showwarning("无题目", "没有符合条件的题目可生成。")
            return

        if self.order_mode.get() == "顺序":
            ordered = []
            for name in self.sheet_names:
                group = [it for it in items if it['topic'] == name]
                ordered.extend(group)
            items = ordered
        else:
            random.shuffle(items)

        questions = []
        answers_parts = []
        for idx, item in enumerate(items, start=1):
            stem = item['stem']
            question_stem = stem.replace('A', '____').replace('B', '____')
            questions.append(f"{idx}. {question_stem}")

            pos_a = stem.find('A')
            pos_b = stem.find('B')
            if pos_a > pos_b:
                pos_a, pos_b = pos_b, pos_a
                item['answer_a'], item['answer_b'] = item['answer_b'], item['answer_a']
            prefix = stem[:pos_a]
            middle = stem[pos_a+1:pos_b]
            suffix = stem[pos_b+1:]
            parts = [
                (f"{idx}. ", False),
                (prefix, False),
                (item['answer_a'], True),
                (middle, False),
                (item['answer_b'], True),
                (suffix, False)
            ]
            answers_parts.append(parts)

        prs = Presentation()
        prs.slide_width = PptCm(33)
        prs.slide_height = PptCm(16)
        blank_layout = prs.slide_layouts[6]
        if blank_layout is None:
            blank_layout = prs.slide_layouts[0]

        def add_textbox_slide(content_list, is_answer=False):
            slide = prs.slides.add_slide(blank_layout)
            for ph in slide.placeholders:
                sp = ph._element
                sp.getparent().remove(sp)
            txBox = slide.shapes.add_textbox(0, 0, PptCm(33), PptCm(16))
            tf = txBox.text_frame
            tf.word_wrap = True

            if is_answer:
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
                            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                        else:
                            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                    p.space_after = PptPt(6)
            else:
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

            if tf.paragraphs[0].text == '':
                p_elem = tf.paragraphs[0]._pPr
                if p_elem is not None:
                    p_elem.getparent().remove(p_elem)

        for i in range(0, len(questions), 5):
            chunk = questions[i:i+5]
            add_textbox_slide(chunk, is_answer=False)

        for i in range(0, len(answers_parts), 5):
            chunk = answers_parts[i:i+5]
            add_textbox_slide(chunk, is_answer=True)

        save_dir = "C:\\"
        filename = f"情境默写_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        full_path = os.path.join(save_dir, filename)
        try:
            prs.save(full_path)
            messagebox.showinfo("成功", f"PPT已保存至：\n{full_path}")
            subprocess.Popen(['explorer', '/select,', full_path])
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
                    subprocess.Popen(['explorer', '/select,', new_path])
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