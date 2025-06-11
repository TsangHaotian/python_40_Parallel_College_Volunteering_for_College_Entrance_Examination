import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
import os

class UniversityInfoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("大学专业录取信息管理")
        self.root.geometry("900x700")

        # 创建带滚动条的主容器
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        
        # 创建画布和滚动条
        canvas = tk.Canvas(container)
        scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        
        # 配置画布
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 放置滚动条和画布
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 创建主框架（放在画布内）
        main_frame = ttk.Frame(canvas, padding="10")
        
        # 将主框架添加到画布中
        canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")

        # 创建左侧的表单区域
        self.form_frame = ttk.LabelFrame(main_frame, text="录入大学专业信息", width=300, height=600)
        self.form_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # 创建右侧的显示区域
        self.display_frame = ttk.LabelFrame(main_frame, text="已录入信息", width=550, height=600)
        self.display_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # 创建底部按钮区域
        self.button_frame = ttk.Frame(main_frame)
        self.button_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        # 设置列和行的权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(0, weight=1)

        # 创建表单控件
        self.create_form()

        # 创建用于显示信息的文本框
        self.info_text = tk.Text(self.display_frame, width=65, height=20)
        self.info_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 创建用于选择删除信息的列表框
        self.info_listbox = tk.Listbox(self.display_frame, width=65, height=10)
        self.info_listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 创建删除按钮
        self.delete_button = ttk.Button(self.display_frame, text="删除选中信息", command=self.delete_info)
        self.delete_button.pack(pady=10)

        
        self.export_button = ttk.Button(self.button_frame, text="导出录取信息", command=self.export_data)
        self.export_button.pack(side=tk.RIGHT, padx=10)

        # 用于存储学校信息的字典
        self.university_info = {}

        # 用于更新列表框的索引
        self.update_listbox()
        
        # 配置画布滚动区域
        def configure_scroll_region(event):
            # 设置画布滚动区域为整个主框架的大小
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 设置画布窗口宽度跟随容器宽度
            canvas.itemconfig(canvas_frame, width=event.width)
            
        # 绑定事件
        main_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_frame, width=e.width))
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_form(self):
        # 学校信息框架
        school_frame = ttk.LabelFrame(self.form_frame, text="学校信息")
        school_frame.pack(fill=tk.X, padx=10, pady=10)

        # 大学代号
        ttk.Label(school_frame, text="学校代号:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.university_code_entry = ttk.Entry(school_frame, width=12)
        self.university_code_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # 大学名称
        ttk.Label(school_frame, text="学校名称:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.university_name_entry = ttk.Entry(school_frame, width=25)
        self.university_name_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 专业组信息框架
        group_frame = ttk.LabelFrame(self.form_frame, text="专业组信息")
        group_frame.pack(fill=tk.X, padx=10, pady=10)

        # 专业组代号
        ttk.Label(group_frame, text="专业组代号:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.group_code_entry = ttk.Entry(group_frame, width=12)
        self.group_code_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # 专业组名称
        ttk.Label(group_frame, text="专业组名称:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.group_name_entry = ttk.Entry(group_frame, width=25)
        self.group_name_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 专业信息框架
        major_frame = ttk.LabelFrame(self.form_frame, text="专业信息")
        major_frame.pack(fill=tk.X, padx=10, pady=10)

        # 专业代号
        ttk.Label(major_frame, text="专业代号:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.major_code_entry = ttk.Entry(major_frame, width=12)
        self.major_code_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # 专业名称
        ttk.Label(major_frame, text="专业名称:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.major_name_entry = ttk.Entry(major_frame, width=25)
        self.major_name_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 录取信息框架
        admission_frame = ttk.LabelFrame(self.form_frame, text="录取信息")
        admission_frame.pack(fill=tk.X, padx=10, pady=10)

        # 专业录取最低分
        ttk.Label(admission_frame, text="录取最低分:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.min_score_entry = ttk.Entry(admission_frame, width=12)
        self.min_score_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # 录取人数
        ttk.Label(admission_frame, text="录取人数:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.admission_count_entry = ttk.Entry(admission_frame, width=12)
        self.admission_count_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 添加按钮
        self.add_button = ttk.Button(self.form_frame, text="添加信息", command=self.add_info)
        self.add_button.pack(pady=10)

        # 清空按钮
        self.clear_button = ttk.Button(self.form_frame, text="清空表单", command=self.clear_form)
        self.clear_button.pack(pady=10)

    def add_info(self):
        university_code = self.university_code_entry.get()
        university_name = self.university_name_entry.get()
        group_code = self.group_code_entry.get()
        group_name = self.group_name_entry.get()
        major_code = self.major_code_entry.get()
        major_name = self.major_name_entry.get()
        min_score = self.min_score_entry.get()
        admission_count = self.admission_count_entry.get()

        # 验证必填字段
        if not university_code or not university_name or not group_code or not major_code or not major_name or not min_score:
            messagebox.showwarning("警告", "请填写所有必要信息！")
            return

        # 检查学校名称是否已存在
        if university_name not in self.university_info:
            self.university_info[university_name] = {
                "代号": university_code,
                "专业组": {},
            }

        # 检查专业组是否已存在
        if group_code not in self.university_info[university_name]["专业组"]:
            self.university_info[university_name]["专业组"][group_code] = {
                "名称": group_name,
                "专业": {}
            }

        # 检查专业名称是否已存在
        if major_code in self.university_info[university_name]["专业组"][group_code]["专业"]:
            messagebox.showwarning("警告", f"{university_name} 的专业组 {group_code} 中已存在专业代号 {major_code}，不能重复录入！")
            return

        # 将信息添加到对应的学校、专业组和专业中
        self.university_info[university_name]["专业组"][group_code]["专业"][major_code] = {
            "名称": major_name,
            "最低分": min_score,
            "录取人数": admission_count if admission_count else "未知"
        }

        # 更新显示区域
        self.update_display()
        self.update_listbox()

        # 清空输入框
        self.clear_form()
        messagebox.showinfo("成功", "信息添加成功！")

    def clear_form(self):
        # 清空所有输入框
        self.university_code_entry.delete(0, tk.END)
        self.university_name_entry.delete(0, tk.END)
        self.group_code_entry.delete(0, tk.END)
        self.group_name_entry.delete(0, tk.END)
        self.major_code_entry.delete(0, tk.END)
        self.major_name_entry.delete(0, tk.END)
        self.min_score_entry.delete(0, tk.END)
        self.admission_count_entry.delete(0, tk.END)

    def update_display(self):
        # 清空文本框
        self.info_text.delete(1.0, tk.END)

        # 按学校名称排序并显示信息
        for university, uni_info in sorted(self.university_info.items()):
            self.info_text.insert(tk.END, f"学校: {university} (代号: {uni_info['代号']})\n")
            
            for group_code, group_info in sorted(uni_info["专业组"].items()):
                self.info_text.insert(tk.END, f"  专业组: {group_code} - {group_info['名称']}\n")
                
                for major_code, major_info in sorted(group_info["专业"].items()):
                    self.info_text.insert(tk.END, f"    专业: {major_code} - {major_info['名称']} | 最低分: {major_info['最低分']} | 录取人数: {major_info['录取人数']}\n")
            
            self.info_text.insert(tk.END, "\n")

    def update_listbox(self):
        # 清空列表框
        self.info_listbox.delete(0, tk.END)

        # 按学校名称排序并显示信息
        for university, uni_info in sorted(self.university_info.items()):
            for group_code, group_info in sorted(uni_info["专业组"].items()):
                for major_code, major_info in sorted(group_info["专业"].items()):
                    display_text = f"{university} - {group_info['名称']} - {major_info['名称']} - {major_info['最低分']}"
                    self.info_listbox.insert(tk.END, display_text)

    def delete_info(self):
        try:
            # 获取选中的信息
            selected_index = self.info_listbox.curselection()[0]
            selected_info = self.info_listbox.get(selected_index)

            # 解析选中的信息
            parts = selected_info.split(" - ")
            university_name = parts[0]
            group_name = parts[1]
            major_name = parts[2]

            # 查找并删除对应的信息
            for group_code, group_info in list(self.university_info[university_name]["专业组"].items()):
                if group_info["名称"] == group_name:
                    for major_code, major_info in list(group_info["专业"].items()):
                        if major_info["名称"] == major_name:
                            del self.university_info[university_name]["专业组"][group_code]["专业"][major_code]
                            
                            # 如果专业组为空，删除专业组
                            if not self.university_info[university_name]["专业组"][group_code]["专业"]:
                                del self.university_info[university_name]["专业组"][group_code]
                            
                            # 如果学校没有专业组，删除学校
                            if not self.university_info[university_name]["专业组"]:
                                del self.university_info[university_name]
                            
                            # 更新显示区域
                            self.update_display()
                            self.update_listbox()
                            messagebox.showinfo("成功", "信息删除成功！")
                            return

        except IndexError:
            messagebox.showwarning("警告", "请选择要删除的信息！")

    def import_zhiyuan_data(self):
        try:
            file_path = filedialog.askopenfilename(title="选择志愿填报表Excel文件", filetypes=[("Excel文件", "*.xlsx")])
            if not file_path:
                return
                
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 检查必要的列是否存在
            required_columns = ["学校代号", "学校名称", "专业组代号", "专业组名称"]
            for i in range(1, 7):
                required_columns.extend([f"专业代号{i}", f"专业名称{i}"])
                
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                messagebox.showerror("错误", f"导入文件缺少必要的列: {', '.join(missing_columns)}")
                return
                
            # 导入数据
            import_count = 0
            for _, row in df.iterrows():
                university_code = row["学校代号"]
                university_name = row["学校名称"]
                group_code = row["专业组代号"]
                group_name = row["专业组名称"]
                
                # 检查学校是否已存在
                if university_name not in self.university_info:
                    self.university_info[university_name] = {
                        "代号": university_code,
                        "专业组": {},
                    }
                
                # 检查专业组是否已存在
                if group_code not in self.university_info[university_name]["专业组"]:
                    self.university_info[university_name]["专业组"][group_code] = {
                        "名称": group_name,
                        "专业": {}
                    }
                
                # 导入专业信息
                for i in range(1, 7):
                    major_code = row.get(f"专业代号{i}")
                    major_name = row.get(f"专业名称{i}")
                    
                    if pd.notna(major_code) and pd.notna(major_name) and major_code and major_name:
                        # 检查专业是否已存在
                        if major_code not in self.university_info[university_name]["专业组"][group_code]["专业"]:
                            self.university_info[university_name]["专业组"][group_code]["专业"][major_code] = {
                                "名称": major_name,
                                "最低分": "未知",
                                "录取人数": "未知"
                            }
                            import_count += 1
            
            # 更新显示
            self.update_display()
            self.update_listbox()
            messagebox.showinfo("成功", f"成功导入 {import_count} 条专业信息！")
            
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def export_data(self):
        try:
            if not self.university_info:
                messagebox.showerror("错误", "没有数据可导出！")
                return

            # 设置默认文件名
            default_filename = "大学录取信息表.xlsx"

            file_path = filedialog.asksaveasfilename(
                title="保存录取信息",
                initialfile=default_filename,  # 设置默认文件名
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )
            if not file_path:
                return

            # 准备导出数据
            export_data = []
            for university, uni_info in self.university_info.items():
                for group_code, group_info in uni_info["专业组"].items():
                    for major_code, major_info in group_info["专业"].items():
                        export_data.append({
                            "学校代号": uni_info["代号"],
                            "学校名称": university,
                            "专业组代号": group_code,
                            "专业组名称": group_info["名称"],
                            "专业代号": major_code,
                            "专业名称": major_info["名称"],
                            "最低录取分": major_info["最低分"],
                            "录取人数": major_info["录取人数"]
                        })

            # 转换为DataFrame并导出
            df = pd.DataFrame(export_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"数据已成功导出到 {file_path}")

        except Exception as e:
            messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversityInfoApp(root)
    root.mainloop()