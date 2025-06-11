import pandas as pd
from tkinter import Tk, Label, Button, Entry, messagebox, filedialog

class VolunteerSimulationSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("志愿填报模拟系统")

        # 初始化数据
        self.volunteer_data = None
        self.university_data = None

        # 创建界面元素
        self.label = Label(root, text="志愿填报模拟系统", font=("Arial", 16))
        self.label.pack(pady=20)

        self.import_volunteer_button = Button(root, text="导入志愿填报表", command=self.import_volunteer_file)
        self.import_volunteer_button.pack(pady=10)

        self.import_university_button = Button(root, text="导入大学录取信息表", command=self.import_university_file)
        self.import_university_button.pack(pady=10)

        self.score_label = Label(root, text="请输入考生分数：")
        self.score_label.pack(pady=5)

        self.score_entry = Entry(root)
        self.score_entry.pack(pady=5)

        self.simulate_button = Button(root, text="模拟填报", command=self.simulate_volunteer)
        self.simulate_button.pack(pady=20)

    def import_volunteer_file(self):
        """导入志愿填报表"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.volunteer_data = pd.read_excel(file_path)
            messagebox.showinfo("成功", "志愿填报表导入成功！")

    def import_university_file(self):
        """导入大学录取信息表"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.university_data = pd.read_excel(file_path)
            messagebox.showinfo("成功", "大学录取信息表导入成功！")

    def simulate_volunteer(self):
        """模拟志愿填报"""
        if self.volunteer_data is None or self.university_data is None:
            messagebox.showerror("错误", "请先导入志愿填报表和大学录取信息表！")
            return

        try:
            student_score = int(self.score_entry.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的考生分数！")
            return

        result = self.process_volunteer(student_score)
        messagebox.showinfo("模拟结果", result)

    def process_volunteer(self, student_score):
        """处理志愿填报逻辑"""
        results = []
        
        for _, row in self.volunteer_data.iterrows():
            university_code = row['学校代号']
            group_code = row['专业组代号']
            major_codes = [row[f'专业代号{i}'] for i in range(1, 7) if pd.notna(row[f'专业代号{i}'])]
            major_names = [row[f'专业名称{i}'] for i in range(1, 7) if pd.notna(row[f'专业名称{i}'])]
            服从调剂 = row['是否服从'] == '是'

            # 查找该专业组的录取信息
            group_info = self.university_data[(self.university_data['学校代号'] == university_code) &
                                              (self.university_data['专业组代号'] == group_code)]

            if group_info.empty:
                results.append(f"未找到{row['学校名称']}的专业组信息，跳过该志愿。")
                continue

            # 获取专业组最低分数线
            group_min_score = group_info['最低录取分'].min()
            
            # 如果分数低于专业组最低分数线，滑档到下一志愿
            if group_min_score > student_score:
                results.append(f"你的分数低于{row['学校名称']}的最低录取分数线{group_min_score}分，滑档。")
                continue
                
            # 检查每个专业是否能录取
            admitted = False
            for major_code, major_name in zip(major_codes, major_names):
                major_info = group_info[group_info['专业代号'] == major_code]
                if not major_info.empty and major_info['最低录取分'].values[0] <= student_score:
                    results.append(f"恭喜！你被{row['学校名称']}的{major_name}专业录取！")
                    return "\n".join(results)  # 录取成功，结束流程
            
            # 如果没有被任何专业录取但分数达到专业组最低分
            if not admitted:
                if 服从调剂:
                    # 找到分数线最低的专业进行调剂
                    lowest_score_major = group_info[group_info['最低录取分'] == group_min_score]
                    results.append(f"你被{row['学校名称']}的{lowest_score_major['专业名称'].values[0]}专业调剂录取！")
                    return "\n".join(results)  # 调剂录取成功，结束流程
                else:
                    # 不服从调剂且所填专业都不符合条件，退档
                    results.append(f"你的分数达到{row['学校名称']}的最低录取分数线，但所填专业均未达到录取条件，且不服从调剂，退档。（请后续等待征集志愿）")
                    # 退档后继续检查下一志愿
                    break;
        
        # 所有志愿都检查完毕，未被录取
        if not results:
            results.append("很遗憾，你未被任何志愿录取，请等待补录或考虑其他院校。")
        
        return "\n".join(results)

if __name__ == "__main__":
    root = Tk()
    app = VolunteerSimulationSystem(root)
    root.mainloop()