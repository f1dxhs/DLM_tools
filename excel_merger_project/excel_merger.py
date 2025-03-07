import pandas as pd
import os
import sys
from tkinter import Tk, filedialog, Button, Label, messagebox, Frame
import tkinter as tk
from tkinter import ttk


class ExcelMerger:
    def __init__(self):
        self.dataframes = []
        self.file_names = []
        
    def read_excel_files(self, file_paths):
        """读取多个Excel文件"""
        self.dataframes = []
        self.file_names = []
        
        for file_path in file_paths:
            try:
                df = pd.read_excel(file_path)
                # 确保必要的列存在
                required_columns = ["图号", "名称", "数量", "单位", "单重（Kg）", "总重（Kg）"]
                for col in required_columns:
                    if col not in df.columns:
                        # 尝试查找列名中包含相似文本的列
                        similar_cols = [c for c in df.columns if col.replace("（", "(").replace("）", ")") in c.replace("（", "(").replace("）", ")")]
                        if similar_cols:
                            df = df.rename(columns={similar_cols[0]: col})
                        else:
                            raise ValueError(f"文件 {os.path.basename(file_path)} 中缺少必要的列: {col}")
                
                # 处理空值图号，确保它们不会被忽略（将NaN转为空字符串）
                df["图号"].fillna("", inplace=True)
                
                self.dataframes.append(df)
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                self.file_names.append(file_name)
            except Exception as e:
                raise Exception(f"读取文件 {file_path} 时出错: {str(e)}")
        
        return len(self.dataframes)
    
    def merge_tables(self):
        """按照指定逻辑合并表格"""
        if not self.dataframes:
            raise Exception("没有数据可合并")
        
        # 创建字典，记录每个标识符的信息
        item_dict = {}
        
        # 用于记录每个文件中的图号顺序
        file_order_maps = []
        
        # 用于存储基础标识符到完整标识符的映射
        base_to_full_identifiers = {}
        
        # 首先遍历所有DataFrame，获取所有唯一项目和它们的顺序
        for idx, df in enumerate(self.dataframes):
            file_name = self.file_names[idx]
            file_order = {}  # 记录当前文件中的图号顺序
            position = 0
            
            for _, row in df.iterrows():
                drawing_no = row["图号"]
                name = row["名称"]
                
                # 获取当前行的备注（如果有）
                current_remark = ""
                if "备注" in row and pd.notna(row["备注"]):
                    current_remark = str(row["备注"]).strip()
                
                # 【修改部分开始】- 根据名称中是否包含"支腿"来决定标识符生成逻辑
                # 检查名称是否包含"支腿"
                has_zhitui = "支腿" in str(name) if pd.notna(name) else False
                
                if pd.isna(drawing_no) or drawing_no == "":
                    # 对于空图号，保持原有逻辑
                    base_identifier = f"__empty__{name}"
                else:
                    if has_zhitui:
                        # 如果包含"支腿"，使用原有的图号+名称作为基础标识符
                        base_identifier = f"{drawing_no}__{name}"
                    else:
                        # 如果不包含"支腿"，只使用图号作为基础标识符
                        base_identifier = f"{drawing_no}"
                
                # 生成完整标识符
                if has_zhitui:
                    # 对于包含"支腿"的项目，使用图号+名称+备注
                    full_identifier = f"{base_identifier}__{current_remark}" if current_remark else base_identifier
                else:
                    # 对于不包含"支腿"的项目，使用图号+备注
                    if current_remark:
                        full_identifier = f"{drawing_no}__{current_remark}"
                    else:
                        full_identifier = f"{drawing_no}"
                
                # 处理空图号的特殊情况
                if pd.isna(drawing_no) or drawing_no == "":
                    full_identifier = base_identifier
                    if current_remark:
                        full_identifier += f"__{current_remark}"
                # 【修改部分结束】
                
                # 记录基础标识符到完整标识符的映射
                if base_identifier not in base_to_full_identifiers:
                    base_to_full_identifiers[base_identifier] = set()
                base_to_full_identifiers[base_identifier].add(full_identifier)
                
                # 记录项目在当前文件中的位置（使用基础标识符）
                if base_identifier not in file_order:
                    file_order[base_identifier] = position
                    position += 1
                
                # 创建项目信息
                if full_identifier not in item_dict:
                    item_dict[full_identifier] = {
                        "图号": drawing_no if not (pd.isna(drawing_no) or drawing_no == "") else "",
                        "名称": name,
                        "数量": 0,
                        "单位": row["单位"],
                        "单重（Kg）": row["单重（Kg）"],
                        "总重（Kg）": row["总重（Kg）"],
                        "files_present": set(),
                        "base_identifier": base_identifier
                    }
                    
                    # 添加备注
                    if current_remark:
                        item_dict[full_identifier]["备注"] = current_remark
                
                # 更新项目信息
                item_dict[full_identifier]["数量"] += row["数量"]
                item_dict[full_identifier]["files_present"].add(idx)
                
                # 设置特定文件的数量
                if f"{file_name}数量" not in item_dict[full_identifier]:
                    item_dict[full_identifier][f"{file_name}数量"] = 0
                item_dict[full_identifier][f"{file_name}数量"] += row["数量"]
            
            # 存储当前文件的顺序映射
            file_order_maps.append(file_order)
        
        # 确保为每个文件设置数量列，如果没有则为0
        for identifier in item_dict:
            for file_name in self.file_names:
                if f"{file_name}数量" not in item_dict[identifier]:
                    item_dict[identifier][f"{file_name}数量"] = 0
        
        # 计算基础标识符的排序
        base_identifiers_order = []
        if len(file_order_maps) > 0:
            first_file_order = file_order_maps[0]
            # 获取第一个文件中的所有基础标识符
            keys_in_first_file = sorted(first_file_order.keys(), key=lambda k: first_file_order[k])
            
            # 创建最终的基础标识符列表，首先添加第一个文件中的所有基础标识符
            base_identifiers_order = list(keys_in_first_file)
            
            # 对于不在第一个文件中的基础标识符，寻找最相似的图号并插入在其后面
            for idx in range(1, len(file_order_maps)):
                file_order = file_order_maps[idx]
                for base_id in file_order.keys():
                    if base_id not in base_identifiers_order:
                        # 提取图号部分进行相似度比较
                        drawing_no = base_id.split("__")[0] if "__" in base_id else base_id
                        
                        # 查找最相似的基础标识符
                        best_match = self._find_similar_base_identifier(base_id, base_identifiers_order)
                        if best_match:
                            # 在最相似的基础标识符后面插入
                            insert_pos = base_identifiers_order.index(best_match) + 1
                            base_identifiers_order.insert(insert_pos, base_id)
                        else:
                            # 如果没有找到相似的，添加到末尾
                            base_identifiers_order.append(base_id)
        
        # 创建最终的完整标识符顺序
        final_order = []
        for base_id in base_identifiers_order:
            # 添加所有具有相同基础标识符的完整标识符
            full_ids = base_to_full_identifiers.get(base_id, set())
            for full_id in sorted(full_ids):
                final_order.append(full_id)
        
        # 创建最终数据集
        result_data = []
        for identifier in final_order:
            item = item_dict[identifier].copy()
            
            # 删除内部使用的辅助字段
            if "files_present" in item:
                del item["files_present"]
            if "base_identifier" in item:
                del item["base_identifier"]
            
            result_data.append(item)
        
        # 创建结果DataFrame
        result_df = pd.DataFrame(result_data)
        
        # 重新排列列
        base_columns = ["图号", "名称", "数量"]
        file_columns = [f"{name}数量" for name in self.file_names]
        other_columns = ["单位", "单重（Kg）", "总重（Kg）"]
        
        # 添加"备注"列如果存在
        if "备注" in result_df.columns:
            other_columns.append("备注")
        
        # 整合所有列
        all_columns = base_columns + file_columns + other_columns
        
        # 只保留DataFrame中实际存在的列
        existing_columns = [col for col in all_columns if col in result_df.columns]
        
        return result_df[existing_columns]
    
    def _find_similar_base_identifier(self, base_id, existing_base_ids):
        """查找与给定基础标识符最相似的标识符"""
        if not existing_base_ids:
            return None
        
        # 如果是空图号，特殊处理
        if base_id.startswith("__empty__"):
            for exist_id in existing_base_ids:
                if exist_id.startswith("__empty__"):
                    return exist_id
            return None
        
        # 提取图号部分
        drawing_no = base_id.split("__")[0] if "__" in base_id else base_id
        
        # 查找最相似的图号
        best_match = None
        highest_similarity = -1
        
        for exist_id in existing_base_ids:
            # 跳过空图号
            if exist_id.startswith("__empty__"):
                continue
            
            # 提取现有标识符的图号部分
            exist_drawing_no = exist_id.split("__")[0] if "__" in exist_id else exist_id
            
            # 计算图号的相似度
            similarity = self._calculate_similarity(drawing_no, exist_drawing_no)
            
            if similarity > highest_similarity:
                highest_similarity = similarity
                best_match = exist_id
        
        # 如果相似度低于阈值，认为没有相似的
        if highest_similarity < 0.3:  # 30%的相似度阈值
            return None
        
        return best_match
    
    def _calculate_similarity(self, str1, str2):
        """计算两个字符串的相似度"""
        # 使用最长公共子序列的长度除以较长字符串的长度作为相似度度量
        # 这种方法偏好相同前缀的字符串
        
        # 找到共同前缀长度
        prefix_len = 0
        for c1, c2 in zip(str1, str2):
            if c1 == c2:
                prefix_len += 1
            else:
                break
        
        # 计算相似度
        max_len = max(len(str1), len(str2))
        if max_len == 0:
            return 0
        
        return prefix_len / max_len
    
    def save_to_excel(self, output_path, merged_df):
        """保存合并后的数据到新的Excel文件"""
        try:
            merged_df.to_excel(output_path, index=False)
            return True
        except Exception as e:
            raise Exception(f"保存文件时出错: {str(e)}")


class ExcelMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel表格合并工具")
        self.root.geometry("600x600")
        self.root.minsize(600, 600)
        
        self.excel_merger = ExcelMerger()
        self.file_paths = []
        
        self.create_widgets()
    
    def create_widgets(self):
        # 创建主框架
        main_frame = Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题标签
        title_label = Label(main_frame, text="Excel表格合并工具", font=("Arial", 16))
        title_label.pack(pady=(0, 20))
        
        # 文件选择按钮
        select_files_button = Button(main_frame, text="选择Excel文件", command=self.select_files, width=20, height=2)
        select_files_button.pack(pady=10)
        
        # 创建一个框架来包含文件列表
        self.files_frame = Frame(main_frame)
        self.files_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建一个Treeview来显示选择的文件
        self.files_treeview = ttk.Treeview(self.files_frame, columns=("文件名", "路径"), show="headings")
        self.files_treeview.heading("文件名", text="文件名")
        self.files_treeview.heading("路径", text="路径")
        self.files_treeview.column("文件名", width=150)
        self.files_treeview.column("路径", width=350)
        self.files_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为Treeview添加滚动条
        scrollbar = ttk.Scrollbar(self.files_frame, orient=tk.VERTICAL, command=self.files_treeview.yview)
        self.files_treeview.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 合并按钮
        merge_button = Button(main_frame, text="合并表格", command=self.merge_files, width=20, height=2)
        merge_button.pack(pady=10)
        
        # 状态标签
        self.status_label = Label(main_frame, text="请选择要合并的Excel文件", fg="gray")
        self.status_label.pack(pady=10)
    
    def select_files(self):
        """选择多个Excel文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        
        if file_paths:
            self.file_paths = file_paths
            self.update_files_treeview()
            self.status_label.config(text=f"已选择 {len(file_paths)} 个文件")
    
    def update_files_treeview(self):
        """更新文件列表显示"""
        # 清除现有项
        for item in self.files_treeview.get_children():
            self.files_treeview.delete(item)
        
        # 添加新文件
        for file_path in self.file_paths:
            file_name = os.path.basename(file_path)
            self.files_treeview.insert("", tk.END, values=(file_name, file_path))
    
    def merge_files(self):
        """合并选择的Excel文件"""
        if not self.file_paths:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            # 读取Excel文件
            num_files = self.excel_merger.read_excel_files(self.file_paths)
            self.status_label.config(text=f"正在合并 {num_files} 个文件...")
            
            # 合并表格
            merged_df = self.excel_merger.merge_tables()
            
            # 选择保存路径
            output_path = filedialog.asksaveasfilename(
                title="保存合并后的Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )
            
            if output_path:
                # 保存到Excel
                self.excel_merger.save_to_excel(output_path, merged_df)
                messagebox.showinfo("成功", f"文件已成功合并并保存至: {output_path}")
                self.status_label.config(text="合并完成")
            else:
                self.status_label.config(text="已取消保存")
        
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.status_label.config(text="合并失败")


def main():
    root = Tk()
    app = ExcelMergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()