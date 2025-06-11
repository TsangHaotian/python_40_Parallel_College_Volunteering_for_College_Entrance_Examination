import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import openpyxl
import os


class ZhiYuanTianBaoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("高考志愿填报系统")
        self.root.geometry("1600x900")  # 增大窗口尺寸以容纳四个志愿表单

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

        # 创建标题
        title_label = ttk.Label(main_frame, text="高考志愿填报表", font=("SimHei", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=10)

        # 创建4个志愿表单
        self.zhiyuan_frames = []
        self.zhiyuan_data = []
        self.zhiyuan_entries = []

        # 定义每个志愿表单的位置（行和列）
        positions = [(1, 0), (1, 2), (2, 0), (2, 2)]

        for i in range(4):
            # 为每个志愿创建数据存储
            data = {
                "志愿序号": i + 1,
                "学校代号": "",
                "学校名称": "",
                "专业组代号": "",
                "专业组名称": "",
                "专业代号1": "",
                "专业名称1": "",
                "专业代号2": "",
                "专业名称2": "",
                "专业代号3": "",
                "专业名称3": "",
                "专业代号4": "",
                "专业名称4": "",
                "专业代号5": "",
                "专业名称5": "",
                "专业代号6": "",
                "专业名称6": "",
                "是否服从": "是",
                "专业调剂": "是"
            }
            self.zhiyuan_data.append(data)

            # 创建志愿框架
            frame = ttk.LabelFrame(main_frame, text=f"第{i + 1}志愿", padding="5")
            row, col = positions[i]
            frame.grid(row=row, column=col, sticky="nsew", padx=20, pady=10)
            self.zhiyuan_frames.append(frame)

            # 创建表单
            entries = self.create_zhiyuan_form(frame, i + 1)
            self.zhiyuan_entries.append(entries)

        # 设置列和行的权重，使其能够均匀分布
        for i in range(4):
            main_frame.columnconfigure(i, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # 创建底部按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=4, pady=20)

        # 导出按钮
        export_button = ttk.Button(button_frame, text="导出所有志愿", command=self.export_all_zhiyuan)
        export_button.pack(side=tk.RIGHT, padx=10)

        # 已保存的志愿列表
        self.saved_zhiyuan = []
        
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

    def create_zhiyuan_form(self, parent, zhiyuan_num):
        # 创建学校信息框架
        school_frame = ttk.LabelFrame(parent, text="学校信息")
        school_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        # 学校代号和名称
        ttk.Label(school_frame, text="学校代号:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        school_code_entry = ttk.Entry(school_frame, width=12)
        school_code_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(school_frame, text="学校名称:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        school_name_entry = ttk.Entry(school_frame, width=25)
        school_name_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 专业组信息
        ttk.Label(school_frame, text="专业组代号:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        major_group_code_entry = ttk.Entry(school_frame, width=12)
        major_group_code_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(school_frame, text="专业组名称:").grid(row=1, column=2, sticky="w", padx=10, pady=5)
        major_group_name_entry = ttk.Entry(school_frame, width=25)
        major_group_name_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        # 创建专业信息框架
        major_frame = ttk.LabelFrame(parent, text="专业信息")
        major_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        # 专业代号和名称（6个）
        major_entries = []
        for i in range(6):
            ttk.Label(major_frame, text=f"专业代号{i + 1}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
            code_entry = ttk.Entry(major_frame, width=12)
            code_entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)

            ttk.Label(major_frame, text=f"专业名称{i + 1}:").grid(row=i, column=2, sticky="w", padx=10, pady=5)
            name_entry = ttk.Entry(major_frame, width=25)
            name_entry.grid(row=i, column=3, sticky="w", padx=5, pady=5)

            major_entries.append((code_entry, name_entry))

        # 创建其他选项框架
        option_frame = ttk.LabelFrame(parent, text="其他选项")
        option_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # 是否服从调剂
        ttk.Label(option_frame, text="是否服从:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        obey_var = tk.StringVar(value="是")
        obey_combo = ttk.Combobox(option_frame, textvariable=obey_var, values=["是", "否"], width=8, state="readonly")
        obey_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # 专业调剂
        ttk.Label(option_frame, text="专业调剂:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        adjust_var = tk.StringVar(value="是")
        adjust_combo = ttk.Combobox(option_frame, textvariable=adjust_var, values=["是", "否"], width=8,
                                    state="readonly")
        adjust_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 按钮框架
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, pady=10)

        # 保存按钮
        save_button = ttk.Button(button_frame, text="保存志愿",
                                 command=lambda z=zhiyuan_num - 1: self.save_zhiyuan(z))
        save_button.grid(row=0, column=0, padx=10)

        # 清空按钮
        clear_button = ttk.Button(button_frame, text="清空表单",
                                  command=lambda z=zhiyuan_num - 1: self.clear_form(z))
        clear_button.grid(row=0, column=1, padx=10)

        # 返回表单条目的引用
        return {
            "school_code": school_code_entry,
            "school_name": school_name_entry,
            "major_group_code": major_group_code_entry,
            "major_group_name": major_group_name_entry,
            "major_entries": major_entries,
            "obey_var": obey_var,
            "adjust_var": adjust_var
        }

    def save_zhiyuan(self, zhiyuan_index):
        # 获取当前志愿的表单条目
        entries = self.zhiyuan_entries[zhiyuan_index]
        data = self.zhiyuan_data[zhiyuan_index]

        # 收集表单数据
        data["学校代号"] = entries["school_code"].get()
        data["学校名称"] = entries["school_name"].get()
        data["专业组代号"] = entries["major_group_code"].get()
        data["专业组名称"] = entries["major_group_name"].get()

        for i, (code_entry, name_entry) in enumerate(entries["major_entries"]):
            data[f"专业代号{i + 1}"] = code_entry.get()
            data[f"专业名称{i + 1}"] = name_entry.get()

        data["是否服从"] = entries["obey_var"].get()
        data["专业调剂"] = entries["adjust_var"].get()

        # 验证必填字段
        if not data["学校代号"] or not data["学校名称"]:
            messagebox.showerror("错误", "学校代号和学校名称为必填项！")
            return

        # 更新或添加到已保存列表
        # 检查是否已存在该志愿序号的数据
        for i, saved in enumerate(self.saved_zhiyuan):
            if saved.get("志愿序号") == data["志愿序号"]:
                self.saved_zhiyuan[i] = data.copy()
                messagebox.showinfo("成功", f"第{data['志愿序号']}志愿信息已更新！")
                return

        # 如果不存在，添加新的
        self.saved_zhiyuan.append(data.copy())
        messagebox.showinfo("成功", f"第{data['志愿序号']}志愿信息已保存！")

    def clear_form(self, zhiyuan_index):
        # 获取当前志愿的表单条目
        entries = self.zhiyuan_entries[zhiyuan_index]

        # 清空所有输入框
        entries["school_code"].delete(0, tk.END)
        entries["school_name"].delete(0, tk.END)
        entries["major_group_code"].delete(0, tk.END)
        entries["major_group_name"].delete(0, tk.END)

        for code_entry, name_entry in entries["major_entries"]:
            code_entry.delete(0, tk.END)
            name_entry.delete(0, tk.END)

        # 重置下拉框
        entries["obey_var"].set("是")
        entries["adjust_var"].set("是")

    def export_all_zhiyuan(self):
        if not self.saved_zhiyuan:
            messagebox.showerror("错误", "没有保存的志愿信息可导出！")
            return

        # 转换为DataFrame
        df = pd.DataFrame(self.saved_zhiyuan)

        # 按志愿序号排序
        df = df.sort_values(by="志愿序号")

        # 导出到Excel
        try:
            file_path = "志愿填报表.xlsx"
            df.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"所有志愿信息已导出到 {os.path.abspath(file_path)}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ZhiYuanTianBaoApp(root)
    root.mainloop()