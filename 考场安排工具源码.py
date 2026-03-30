import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill


class ExamArrangeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("沂水县第二实验中学考场安排工具")
        self.root.geometry("800x650")  # 微调高度适配删除行后的布局
        self.root.resizable(True, True)

        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('微软雅黑', 9))
        style.configure('TButton', font=('微软雅黑', 8))
        style.configure('TRadiobutton', font=('微软雅黑', 9))
        style.configure('TLabelframe.Label', font=('微软雅黑', 9, 'bold'))

        # 定义文件夹结构
        self.desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.main_folder = os.path.join(self.desktop, "考场打印材料")
        self.input_folder = os.path.join(self.main_folder, "输入文件")
        self.output_folder = os.path.join(self.main_folder, "输出结果")

        # 创建文件夹
        for folder in [self.main_folder, self.input_folder, self.output_folder]:
            if not os.path.exists(folder):
                os.makedirs(folder)

        # 固定输出文件完整路径
        self.student_output_file = os.path.join(self.input_folder, "学生考试信息.xlsx")
        self.capacity_output_file = os.path.join(self.output_folder, "考场容量.xlsx")

        # 模板路径
        self.score_template_path = os.path.join(self.input_folder, "学生原始成绩.xlsx")
        self.default_capacity_template_path = os.path.join(self.input_folder, "考场容量模板.xlsx")

        # 生成模板
        self.init_templates()

        # 变量
        self.input_file = tk.StringVar(value=self.score_template_path)
        self.capacity_template_file = tk.StringVar(value=self.default_capacity_template_path)
        # 输出文件完整路径显示变量
        self.student_output_path_var = tk.StringVar(value=self.student_output_file)
        self.capacity_output_path_var = tk.StringVar(value=self.capacity_output_file)
        self.exam_prefix = tk.StringVar(value="701020")

        # 存储考场行控件
        self.capacity_rows = []  # 每项: (frame, entry, delete_btn, court_num)

        # 创建界面
        self.create_widgets()

        # 绑定路径变量监听，实现实时联动校验
        self.capacity_template_file.trace_add('write', self.check_capacity_path)

        # 默认添加9个考场
        for i in range(1, 10):
            self.add_capacity_row(i)

    def check_capacity_path(self, *args):
        """联动核心：实时校验容量模板路径，保证导入导出永远使用当前输入框的路径"""
        current_path = self.capacity_template_file.get().strip()
        # 清空输入框时恢复默认路径
        if not current_path:
            self.capacity_template_file.set(self.default_capacity_template_path)

    def init_templates(self):
        """生成成绩模板和考场容量模板（如果不存在）"""
        # 学生原始成绩模板
        if not os.path.exists(self.score_template_path):
            df_score = pd.DataFrame({
                "姓名": ["示例学生1", "示例学生2"],
                "班级": [1, 1],
                "总分": [650, 620],
                "年级排名": [10, 25]
            })
            with pd.ExcelWriter(self.score_template_path, engine='openpyxl') as writer:
                df_score.to_excel(writer, index=False, sheet_name="成绩")
                worksheet = writer.sheets["成绩"]
                worksheet.cell(row=1, column=5, value="提示：总分和年级排名至少填写一项，系统将自动识别。")
                worksheet.cell(row=2, column=5, value="（请勿删除或修改列名，班级必须为数字）")

        # 考场容量模板（使用默认路径初始化）
        if not os.path.exists(self.default_capacity_template_path):
            default_court_count = 9
            df_cap = pd.DataFrame({
                "考场": list(range(1, default_court_count + 1)),
                "容量": [50] * default_court_count
            })
            with pd.ExcelWriter(self.default_capacity_template_path, engine='openpyxl') as writer:
                df_cap.to_excel(writer, index=False, sheet_name="容量")
                worksheet = writer.sheets["容量"]

                # 表头基础提示
                worksheet.cell(row=1, column=3, value="⚠️  核心规则说明")
                worksheet.cell(row=2, column=3, value="1. 请在此填写前N个考场的固定容量")
                worksheet.cell(row=3, column=3, value="2. 最终的最后一个考场，无需在此填写，程序将根据学生总数自动计算")
                worksheet.cell(row=4, column=3, value="3. 请勿修改「考场」「容量」列名，仅修改容量数值即可")

                # 最后一行考场的醒目提示
                last_data_row = default_court_count + 1
                tip_cell = worksheet.cell(row=last_data_row, column=3,
                                          value="★ 模板手动考场到此结束，无需往下新增行，最后一个考场自动计算")
                tip_cell.font = Font(color="FF0000", bold=True)
                tip_cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件设置（已删除输出根目录行，其余结构完全不变）
        file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding="8")
        file_frame.pack(fill=tk.X, pady=(0, 8))

        # 第一行：成绩文件
        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="成绩文件:", width=14).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.input_file, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览", command=self.browse_input).pack(side=tk.LEFT, padx=5)

        # 第二行：容量模板
        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="容量模板:", width=14).pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.capacity_template_file, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X,
                                                                                 expand=True)
        ttk.Button(row2, text="浏览", command=self.browse_capacity_template).pack(side=tk.LEFT, padx=5)

        # 第三行：学生考试信息完整路径
        row3 = ttk.Frame(file_frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="学生信息输出:", width=14, foreground="#0066CC").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.student_output_path_var, state="readonly", width=60).pack(side=tk.LEFT,
                                                                                                    padx=5, fill=tk.X,
                                                                                                    expand=True)

        # 第四行：考场容量完整路径
        row4 = ttk.Frame(file_frame)
        row4.pack(fill=tk.X, pady=3)
        ttk.Label(row4, text="考场容量输出:", width=14, foreground="#0066CC").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.capacity_output_path_var, state="readonly", width=60).pack(side=tk.LEFT,
                                                                                                     padx=5, fill=tk.X,
                                                                                                     expand=True)

        # 基本信息
        info_frame = ttk.LabelFrame(main_frame, text="基本信息", padding="8")
        info_frame.pack(fill=tk.X, pady=(0, 8))

        row5 = ttk.Frame(info_frame)
        row5.pack(fill=tk.X, pady=3)
        ttk.Label(row5, text="考号前缀:", width=14).pack(side=tk.LEFT)
        ttk.Entry(row5, textvariable=self.exam_prefix, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(row5, text=" (例:701020 → 701020001)").pack(side=tk.LEFT, padx=5)

        # 排序依据提示
        row6 = ttk.Frame(info_frame)
        row6.pack(fill=tk.X, pady=3)
        ttk.Label(row6, text="排序规则:", width=14).pack(side=tk.LEFT)
        ttk.Label(row6, text="系统自动检测：优先按“总分”降序，若无总分则按“年级排名”升序", foreground="gray").pack(
            side=tk.LEFT, padx=5)

        # 考场容量设置
        capacity_frame = ttk.LabelFrame(main_frame,
                                        text="考场容量设置 | 仅需填写前N个考场，最后一个考场自动计算，无需手动添加",
                                        padding="8")
        capacity_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # 按钮栏
        btn_bar = ttk.Frame(capacity_frame)
        btn_bar.pack(fill=tk.X, pady=3)
        ttk.Button(btn_bar, text="添加考场", command=self.add_capacity_row).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="从模板导入", command=self.import_capacity_from_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="导出当前为模板", command=self.export_capacity_template).pack(side=tk.LEFT, padx=3)

        # 滚动区域（恢复高度适配布局）
        canvas = tk.Canvas(capacity_frame, borderwidth=0, highlightthickness=0, height=180)
        scrollbar = ttk.Scrollbar(capacity_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.capacity_container = self.scrollable_frame

        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=5)
        ttk.Button(bottom_frame, text="重置成绩模板", command=self.reset_score_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="开始生成安排", command=self.generate_arrangement).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="退出", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

    def browse_input(self):
        filename = filedialog.askopenfilename(initialdir=self.input_folder, filetypes=[("Excel文件", "*.xlsx *.xls")])
        if filename:
            self.input_file.set(filename)

    def browse_capacity_template(self):
        """联动优化：浏览时默认打开当前输入框路径的目录，选择后自动更新输入框"""
        # 获取当前输入框里的路径所在目录，保持联动
        current_path = self.capacity_template_file.get()
        initial_dir = os.path.dirname(current_path) if os.path.exists(
            os.path.dirname(current_path)) else self.input_folder

        filename = filedialog.askopenfilename(
            initialdir=initial_dir,
            filetypes=[("Excel文件", "*.xlsx *.xls")],
            title="选择容量模板文件"
        )
        if filename:
            self.capacity_template_file.set(filename)  # 自动更新输入框，导入导出同步生效

    def add_capacity_row(self, court_num=None):
        """添加一个考场输入行"""
        if court_num is None:
            court_num = len(self.capacity_rows) + 1
        row_frame = ttk.Frame(self.capacity_container)
        row_frame.pack(fill=tk.X, pady=2)
        lbl = ttk.Label(row_frame, text=f"考场 {court_num}:", width=8)
        lbl.pack(side=tk.LEFT)
        entry = ttk.Entry(row_frame, width=10)
        entry.insert(0, "50")
        entry.pack(side=tk.LEFT, padx=5)
        delete_btn = ttk.Button(row_frame, text="删除", command=lambda: self.delete_capacity_row(row_frame, court_num))
        delete_btn.pack(side=tk.LEFT, padx=5)
        self.capacity_rows.append((row_frame, entry, delete_btn, court_num))

    def delete_capacity_row(self, row_frame, court_num):
        """删除考场行"""
        for i, (f, e, btn, num) in enumerate(self.capacity_rows):
            if f == row_frame:
                f.destroy()
                self.capacity_rows.pop(i)
                self.renumber_courts()
                break

    def renumber_courts(self):
        """重新编号考场行"""
        for idx, (frame, entry, btn, _) in enumerate(self.capacity_rows):
            new_num = idx + 1
            for child in frame.winfo_children():
                if isinstance(child, ttk.Label) and child.cget('text').startswith('考场'):
                    child.config(text=f"考场 {new_num}:")
                    break
            self.capacity_rows[idx] = (frame, entry, btn, new_num)

    def get_capacities(self):
        """获取所有手动设置的考场容量"""
        caps = []
        for frame, entry, btn, num in self.capacity_rows:
            try:
                val = int(entry.get())
                if val <= 0:
                    raise ValueError
                caps.append(val)
            except:
                messagebox.showerror("错误", f"考场 {num} 的容量必须为正整数")
                return None
        return caps

    def import_capacity_from_template(self):
        """联动核心：完全从输入框当前路径导入，无任何硬编码路径"""
        template_path = self.capacity_template_file.get().strip()

        # 多层校验，完全联动输入框内容
        if not template_path:
            messagebox.showerror("错误", "请先选择容量模板文件，或在输入框填写有效路径")
            return
        if not os.path.exists(template_path):
            messagebox.showerror("错误", f"模板文件不存在：\n{template_path}\n请检查路径是否正确")
            return
        if not template_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror("错误", "仅支持.xlsx/.xls格式的Excel文件")
            return

        try:
            # 只读取有效数据，跳过空行和提示行
            df = pd.read_excel(template_path, sheet_name="容量")
            df = df.dropna(subset=['容量'])
            df['容量'] = pd.to_numeric(df['容量'], errors='coerce')
            df = df.dropna(subset=['容量'])
            caps = df['容量'].astype(int).tolist()

            if not caps:
                messagebox.showerror("错误", "模板中未读取到有效的考场容量数据")
                return

            # 清除现有行
            for frame, _, _, _ in self.capacity_rows:
                frame.destroy()
            self.capacity_rows.clear()
            # 添加新行
            for i, cap in enumerate(caps):
                self.add_capacity_row(i + 1)
                self.capacity_rows[-1][1].delete(0, tk.END)
                self.capacity_rows[-1][1].insert(0, str(cap))
            messagebox.showinfo("成功",
                                f"已从以下路径导入 {len(caps)} 个有效考场：\n{template_path}\n最后一个考场将自动计算")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败：{str(e)}")

    def export_capacity_template(self):
        """联动核心：完全导出到输入框当前路径，无任何硬编码路径"""
        export_path = self.capacity_template_file.get().strip()

        # 路径校验，完全联动输入框内容
        if not export_path:
            messagebox.showerror("错误", "请先在输入框填写导出路径，或通过浏览选择目标文件")
            return
        if not export_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror("错误", "仅支持导出为.xlsx/.xls格式的Excel文件")
            return

        caps = self.get_capacities()
        if caps is None:
            return

        # 自动创建目标文件夹（如果不存在）
        target_dir = os.path.dirname(export_path)
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

        # 覆盖文件前确认
        if os.path.exists(export_path):
            if not messagebox.askyesno("确认覆盖", f"目标文件已存在：\n{export_path}\n是否覆盖？"):
                return

        court_count = len(caps)
        df = pd.DataFrame({
            "考场": list(range(1, court_count + 1)),
            "容量": caps
        })

        try:
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="容量")
                worksheet = writer.sheets["容量"]

                # 表头基础提示
                worksheet.cell(row=1, column=3, value="⚠️  核心规则说明")
                worksheet.cell(row=2, column=3, value="1. 请在此填写前N个考场的固定容量")
                worksheet.cell(row=3, column=3, value="2. 最终的最后一个考场，无需在此填写，程序将根据学生总数自动计算")
                worksheet.cell(row=4, column=3, value="3. 请勿修改「考场」「容量」列名，仅修改容量数值即可")

                # 最后一行考场的醒目提示
                last_data_row = court_count + 1
                tip_cell = worksheet.cell(row=last_data_row, column=3,
                                          value="★ 模板手动考场到此结束，无需往下新增行，最后一个考场自动计算")
                tip_cell.font = Font(color="FF0000", bold=True)
                tip_cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

            # 导出成功后不修改路径，保持联动，提示当前导出路径
            messagebox.showinfo("成功", f"当前考场容量已导出至：\n{export_path}\n导入/导出功能已同步更新此路径")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def reset_score_template(self):
        """重置学生原始成绩模板"""
        try:
            df_score = pd.DataFrame({
                "姓名": ["示例学生1", "示例学生2"],
                "班级": [1, 1],
                "总分": [650, 620],
                "年级排名": [10, 25]
            })
            with pd.ExcelWriter(self.score_template_path, engine='openpyxl') as writer:
                df_score.to_excel(writer, index=False, sheet_name="成绩")
                worksheet = writer.sheets["成绩"]
                worksheet.cell(row=1, column=5, value="提示：总分和年级排名至少填写一项，系统将自动识别。")
                worksheet.cell(row=2, column=5, value="（请勿删除或修改列名，班级必须为数字）")
            messagebox.showinfo("成功", f"成绩模板已重置：\n{self.score_template_path}")
        except Exception as e:
            messagebox.showerror("错误", f"重置失败：{e}")

    def generate_arrangement(self):
        """核心生成逻辑（完全不变）"""
        source_file_path = self.input_file.get()

        if not os.path.exists(source_file_path):
            messagebox.showerror("错误", "成绩文件不存在，请检查路径。")
            return

        caps = self.get_capacities()
        if caps is None or len(caps) == 0:
            messagebox.showerror("错误", "请至少设置一个有效考场容量")
            return

        try:
            # 固定文件保存路径，与UI显示完全一致
            output_student_file = self.student_output_file
            output_capacity_file = self.capacity_output_file

            df = pd.read_excel(source_file_path)

            # 自动判断排序依据
            rank_col = None
            ascending_rank = True

            has_total = '总分' in df.columns
            has_rank = '年级排名' in df.columns

            if not has_total and not has_rank:
                messagebox.showerror("错误", "成绩文件中必须包含“总分”或“年级排名”列！")
                return

            if has_total:
                rank_col = '总分'
                ascending_rank = False  # 总分降序
                sort_msg = "总分（从高到低）"
            else:
                rank_col = '年级排名'
                ascending_rank = True  # 排名升序
                sort_msg = "年级排名（从先到后）"

            # 检查必备列
            required = ['姓名', '班级', rank_col]
            for col in required:
                if col not in df.columns:
                    messagebox.showerror("错误", f"成绩文件中缺少必填列：{col}")
                    return

            # 数据清洗：去除空行
            df = df.dropna(subset=['姓名', '班级'])
            total_students = len(df)
            if total_students == 0:
                messagebox.showerror("错误", "成绩文件中没有有效的学生数据")
                return

            # 排序：先按班级内排名，再S型分班
            df = df.sort_values(['班级', rank_col, '姓名'], ascending=[True, ascending_rank, True])
            df['班级内序号'] = df.groupby('班级').cumcount() + 1
            df_sorted = df.sort_values(['班级内序号', '班级'], ascending=[True, True])

            # 核心修复：仅当有剩余学生时，才添加自动计算考场
            fixed_capacity = sum(caps)
            last_court_capacity = total_students - fixed_capacity

            if last_court_capacity < 0:
                messagebox.showerror("错误",
                                     f"前{len(caps)}个考场总容量({fixed_capacity})，已超过学生总数({total_students})！\n请减少考场数量或降低单考场容量。")
                return

            all_caps = caps.copy()
            has_auto_court = False
            if last_court_capacity > 0:
                all_caps.append(last_court_capacity)
                has_auto_court = True

            # 分配考场和座号
            allocations = []
            court_idx = 0
            seat_in_court = 0
            for idx, (_, student) in enumerate(df_sorted.iterrows()):
                if seat_in_court >= all_caps[court_idx]:
                    court_idx += 1
                    seat_in_court = 0
                exam_id = f"{self.exam_prefix.get()}{(idx + 1):03d}"
                allocations.append({
                    '准考证号': exam_id,
                    '姓名': student['姓名'],
                    '班级': int(student['班级']),
                    '考场号': court_idx + 1,
                    '座号': seat_in_court + 1
                })
                seat_in_court += 1

            # 按考场+座号排序后导出
            result_df = pd.DataFrame(allocations)
            result_df = result_df.sort_values(['考场号', '座号'])
            result_df.to_excel(output_student_file, index=False)

            # 生成考场容量汇总表
            court_info = []
            manual_court_count = len(caps)
            for court_num in range(1, len(all_caps) + 1):
                court_students = result_df[result_df['考场号'] == court_num]
                student_count = len(court_students)
                start_id = court_students.iloc[0]['准考证号'] if student_count > 0 else ''
                end_id = court_students.iloc[-1]['准考证号'] if student_count > 0 else ''
                set_cap = all_caps[court_num - 1] if court_num <= manual_court_count else '自动计算'
                court_info.append({
                    '考场号': court_num,
                    '设置容量': set_cap,
                    '实际人数': student_count,
                    '起始考号': start_id,
                    '截止考号': end_id
                })
            capacity_df = pd.DataFrame(court_info)
            capacity_df.to_excel(output_capacity_file, index=False)

            # 优化成功提示文案，显示完整路径
            auto_court_tip = f"（含1个自动计算的收尾考场，容量{last_court_capacity}人）" if has_auto_court else "（无自动收尾考场，手动考场已容纳全部学生）"
            messagebox.showinfo("生成成功",
                                f"考场安排已完成！\n\n排序依据：{sort_msg}\n学生总数：{total_students}人\n考场总数：{len(all_caps)}个{auto_court_tip}\n\n文件已保存至：\n- 学生考试信息：{output_student_file}\n- 考场容量：{output_capacity_file}")
        except Exception as e:
            messagebox.showerror("生成失败", f"程序运行出错：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExamArrangeApp(root)
    root.mainloop()