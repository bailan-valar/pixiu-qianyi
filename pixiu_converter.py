import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from pathlib import Path

class PixiuConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("貔貅记账转换工具")
        self.root.geometry("1200x800")

        self.source_file = None
        self.source_data = None
        self.category_mapping = {}
        self.account_mapping = {}

        # 加载已保存的映射配置
        self.load_mappings()

        # 创建主界面
        self.create_widgets()

    def load_mappings(self):
        """加载已保存的映射配置"""
        if os.path.exists('category_mapping.json'):
            with open('category_mapping.json', 'r', encoding='utf-8') as f:
                self.category_mapping = json.load(f)
        if os.path.exists('account_mapping.json'):
            with open('account_mapping.json', 'r', encoding='utf-8') as f:
                self.account_mapping = json.load(f)

    def save_mappings(self):
        """保存映射配置"""
        with open('category_mapping.json', 'w', encoding='utf-8') as f:
            json.dump(self.category_mapping, f, ensure_ascii=False, indent=2)
        with open('account_mapping.json', 'w', encoding='utf-8') as f:
            json.dump(self.account_mapping, f, ensure_ascii=False, indent=2)

    def create_widgets(self):
        """创建界面组件"""
        # 创建Notebook
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 文件选择标签页
        file_frame = ttk.Frame(notebook)
        notebook.add(file_frame, text='1. 文件选择')
        self.create_file_selection_tab(file_frame)

        # 分类映射标签页
        category_frame = ttk.Frame(notebook)
        notebook.add(category_frame, text='2. 分类映射')
        self.create_category_mapping_tab(category_frame)

        # 账户映射标签页
        account_frame = ttk.Frame(notebook)
        notebook.add(account_frame, text='3. 账户映射')
        self.create_account_mapping_tab(account_frame)

        # 导出标签页
        export_frame = ttk.Frame(notebook)
        notebook.add(export_frame, text='4. 导出')
        self.create_export_tab(export_frame)

    def create_file_selection_tab(self, parent):
        """创建文件选择标签页"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=20, pady=20)

        # 文件选择
        ttk.Label(frame, text='选择源文件 (Result_2.csv 格式):', font=('Arial', 12, 'bold')).pack(anchor='w', pady=10)

        file_select_frame = ttk.Frame(frame)
        file_select_frame.pack(fill='x', pady=5)

        self.file_path_label = ttk.Label(file_select_frame, text='未选择文件', relief='sunken', anchor='w')
        self.file_path_label.pack(side='left', fill='x', expand=True, padx=(0, 10))

        ttk.Button(file_select_frame, text='浏览...', command=self.select_source_file).pack(side='right')

        # 文件信息显示
        info_frame = ttk.LabelFrame(frame, text='文件信息')
        info_frame.pack(fill='both', expand=True, pady=20)

        self.file_info_text = tk.Text(info_frame, height=15, wrap='word')
        self.file_info_text.pack(fill='both', expand=True, padx=5, pady=5)

        # 加载貔貅模板
        template_frame = ttk.Frame(frame)
        template_frame.pack(fill='x', pady=10)

        ttk.Label(template_frame, text='貔貅模板文件:').pack(side='left')
        self.template_label = ttk.Label(template_frame, text='貔貅记账#日常收支交易导入模版.csv', foreground='green')
        self.template_label.pack(side='left', padx=10)

        if os.path.exists('貔貅记账#日常收支交易导入模版.csv'):
            ttk.Label(template_frame, text='✓ 已找到', foreground='green').pack(side='left')
        else:
            ttk.Label(template_frame, text='✗ 未找到', foreground='red').pack(side='left')

    def create_category_mapping_tab(self, parent):
        """创建分类映射标签页"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=20, pady=20)

        # 控制面板
        control_frame = ttk.Frame(frame)
        control_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(control_frame, text='分类映射配置', font=('Arial', 12, 'bold')).pack(side='left')
        ttk.Button(control_frame, text='保存映射', command=self.save_mappings).pack(side='right', padx=5)
        ttk.Button(control_frame, text='加载映射', command=self.load_mappings).pack(side='right', padx=5)
        ttk.Button(control_frame, text='自动匹配', command=self.auto_match_categories).pack(side='right', padx=5)

        # 分割窗口
        paned = ttk.PanedWindow(frame, orient='horizontal')
        paned.pack(fill='both', expand=True)

        # 左侧：源分类列表
        left_frame = ttk.LabelFrame(paned, text='源分类 (从 Result_2.csv)')
        paned.add(left_frame, weight=1)

        # 创建Treeview显示源分类（添加状态列）
        columns = ('income_type', 'category', 'sub_category', 'status')
        self.source_tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=20)
        self.source_tree.heading('income_type', text='收支类型')
        self.source_tree.heading('category', text='分类')
        self.source_tree.heading('sub_category', text='子分类')
        self.source_tree.heading('status', text='状态')

        self.source_tree.column('income_type', width=80)
        self.source_tree.column('category', width=150)
        self.source_tree.column('sub_category', width=150)
        self.source_tree.column('status', width=80)

        # 定义状态标签颜色
        self.source_tree.tag_configure('mapped', foreground='green')
        self.source_tree.tag_configure('deleted', foreground='red')
        self.source_tree.tag_configure('unmapped', foreground='gray')

        scrollbar_left = ttk.Scrollbar(left_frame, orient='vertical', command=self.source_tree.yview)
        self.source_tree.configure(yscrollcommand=scrollbar_left.set)

        self.source_tree.pack(side='left', fill='both', expand=True)
        scrollbar_left.pack(side='right', fill='y')

        self.source_tree.bind('<<TreeviewSelect>>', self.on_source_category_select)

        # 右侧：目标分类选择
        right_frame = ttk.LabelFrame(paned, text='目标分类 (貔貅记账)')
        paned.add(right_frame, weight=1)

        # 收支类型选择
        type_frame = ttk.Frame(right_frame)
        type_frame.pack(fill='x', pady=5)

        ttk.Label(type_frame, text='收支类型:').pack(side='left', padx=5)
        self.income_type_var = tk.StringVar(value='支出')
        type_combo = ttk.Combobox(type_frame, textvariable=self.income_type_var, values=['收入', '支出'], state='readonly', width=10)
        type_combo.pack(side='left', padx=5)
        type_combo.bind('<<ComboboxSelected>>', self.on_income_type_change)

        # 搜索框
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill='x', pady=5)

        ttk.Label(search_frame, text='搜索分类:').pack(side='left', padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind('<KeyRelease>', self.on_search_category)

        # 目标分类选择
        target_frame = ttk.Frame(right_frame)
        target_frame.pack(fill='both', expand=True, pady=5)

        ttk.Label(target_frame, text='选择目标分类:').pack(anchor='w', padx=5)

        # 目标分类列表
        self.target_listbox = tk.Listbox(target_frame, height=15)
        target_scrollbar = ttk.Scrollbar(target_frame, orient='vertical', command=self.target_listbox.yview)
        self.target_listbox.configure(yscrollcommand=target_scrollbar.set)

        self.target_listbox.pack(side='left', fill='both', expand=True, padx=(5, 0), pady=5)
        target_scrollbar.pack(side='left', fill='y', pady=5)

        # 当前映射显示
        mapping_frame = ttk.LabelFrame(right_frame, text='当前选中分类的映射')
        mapping_frame.pack(fill='x', pady=10, padx=5)

        self.current_mapping_label = ttk.Label(mapping_frame, text='未选择', foreground='gray')
        self.current_mapping_label.pack(padx=10, pady=5)

        # 设置映射按钮
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill='x', pady=5)

        ttk.Button(btn_frame, text='设置映射', command=self.set_category_mapping).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='标记为删除', command=self.set_category_delete).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='清除映射', command=self.clear_category_mapping).pack(side='left', padx=5)

        # 加载目标分类
        self.load_target_categories()

    def create_account_mapping_tab(self, parent):
        """创建账户映射标签页"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=20, pady=20)

        # 控制面板
        control_frame = ttk.Frame(frame)
        control_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(control_frame, text='账户映射配置', font=('Arial', 12, 'bold')).pack(side='left')
        ttk.Button(control_frame, text='保存映射', command=self.save_mappings).pack(side='right', padx=5)
        ttk.Button(control_frame, text='加载映射', command=self.load_mappings).pack(side='right', padx=5)

        # 分割窗口
        paned = ttk.PanedWindow(frame, orient='horizontal')
        paned.pack(fill='both', expand=True)

        # 左侧：源账户列表
        left_frame = ttk.LabelFrame(paned, text='源账户')
        paned.add(left_frame, weight=1)

        self.source_account_listbox = tk.Listbox(left_frame, height=25)
        source_account_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.source_account_listbox.yview)
        self.source_account_listbox.configure(yscrollcommand=source_account_scrollbar.set)

        self.source_account_listbox.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        source_account_scrollbar.pack(side='right', fill='y', pady=5)

        self.source_account_listbox.bind('<<ListboxSelect>>', self.on_source_account_select)

        # 右侧：目标账户输入
        right_frame = ttk.LabelFrame(paned, text='目标账户设置')
        paned.add(right_frame, weight=1)

        # 账户映射设置
        setting_frame = ttk.Frame(right_frame)
        setting_frame.pack(fill='x', pady=10)

        ttk.Label(setting_frame, text='目标账户名称:').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.target_account_entry = ttk.Entry(setting_frame, width=30)
        self.target_account_entry.grid(row=0, column=1, padx=5, pady=5)

        # 快速选择
        quick_frame = ttk.LabelFrame(right_frame, text='快速选择常用账户')
        quick_frame.pack(fill='x', pady=10, padx=5)

        quick_accounts = ['微信零钱', '银行卡', '支付宝', '现金', '信用卡']
        for i, account in enumerate(quick_accounts):
            ttk.Button(quick_frame, text=account,
                      command=lambda a=account: self.set_target_account(a)).grid(row=i//3, column=i%3, padx=5, pady=5)

        # 当前映射显示
        current_mapping_frame = ttk.LabelFrame(right_frame, text='当前选中账户的映射')
        current_mapping_frame.pack(fill='x', pady=10, padx=5)

        self.current_account_mapping_label = ttk.Label(current_mapping_frame, text='未选择', foreground='gray')
        self.current_account_mapping_label.pack(padx=10, pady=5)

        # 设置映射按钮
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill='x', pady=5)

        ttk.Button(btn_frame, text='设置映射', command=self.set_account_mapping).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='清除映射', command=self.clear_account_mapping).pack(side='left', padx=5)

    def create_export_tab(self, parent):
        """创建导出标签页"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=20, pady=20)

        ttk.Label(frame, text='导出设置', font=('Arial', 12, 'bold')).pack(pady=10)

        # 导出选项
        options_frame = ttk.LabelFrame(frame, text='导出选项')
        options_frame.pack(fill='x', pady=10)

        self.export_unmapped_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text='导出未映射的记录（使用原分类名称）',
                       variable=self.export_unmapped_var).pack(anchor='w', padx=10, pady=5)

        # 预览区域
        preview_frame = ttk.LabelFrame(frame, text='数据预览')
        preview_frame.pack(fill='both', expand=True, pady=10)

        self.preview_text = tk.Text(preview_frame, height=20, wrap='word')
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)

        self.preview_text.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        preview_scrollbar.pack(side='right', fill='y', pady=5)

        # 导出按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=10)

        ttk.Button(btn_frame, text='预览转换结果', command=self.preview_conversion).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='导出CSV', command=self.export_csv).pack(side='left', padx=5)

    def select_source_file(self):
        """选择源文件"""
        file_path = filedialog.askopenfilename(
            title='选择 Result_2.csv 文件',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )

        if file_path:
            self.source_file = file_path
            self.file_path_label.config(text=file_path)
            self.load_source_data()

    def load_source_data(self):
        """加载源数据"""
        try:
            self.source_data = pd.read_csv(self.source_file)

            # 显示文件信息
            info = f"文件: {self.source_file}\n"
            info += f"总记录数: {len(self.source_data)}\n"
            info += f"列数: {len(self.source_data.columns)}\n\n"
            info += "列名:\n"
            for col in self.source_data.columns:
                info += f"  - {col}\n"

            self.file_info_text.delete('1.0', tk.END)
            self.file_info_text.insert('1.0', info)

            # 加载源分类到Treeview
            self.load_source_categories()
            self.load_source_accounts()

            messagebox.showinfo('成功', '文件加载成功！')

        except Exception as e:
            messagebox.showerror('错误', f'加载文件失败: {str(e)}')

    def load_source_categories(self):
        """加载源分类"""
        # 清空Treeview
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)

        if self.source_data is None:
            return

        # 获取唯一分类组合
        categories = self.source_data[['income_type', 'category', 'sub_category']].drop_duplicates()

        for _, row in categories.iterrows():
            income_type = row['income_type']
            category = row['category']
            sub_category = row['sub_category']

            # 创建映射键
            key = f"{income_type}|{category}|{sub_category}"

            # 检查映射状态
            if key in self.category_mapping:
                if self.category_mapping[key] == '__DELETE__':
                    status = '已删除'
                    tag = 'deleted'
                else:
                    status = '已映射'
                    tag = 'mapped'
            else:
                status = '未映射'
                tag = 'unmapped'

            # 插入数据，包含状态列
            self.source_tree.insert('', 'end', iid=key, values=(income_type, category, sub_category, status), tags=(tag,))

    def load_source_accounts(self):
        """加载源账户"""
        self.source_account_listbox.delete('0', tk.END)

        if self.source_data is None:
            return

        # 获取唯一账户
        if 'account_name' in self.source_data.columns:
            accounts = self.source_data['account_name'].dropna().unique()
            for account in sorted(accounts):
                self.source_account_listbox.insert('end', account)

    def load_target_categories(self):
        """加载目标分类"""
        if not os.path.exists('target_category.json'):
            return

        with open('target_category.json', 'r', encoding='utf-8') as f:
            target_categories = json.load(f)

        self.target_categories_data = target_categories
        self.update_target_listbox()

    def update_target_listbox(self):
        """更新目标分类列表"""
        self.target_listbox.delete('0', 'end')

        income_type = self.income_type_var.get()
        if income_type in self.target_categories_data:
            for category in self.target_categories_data[income_type]:
                self.target_listbox.insert('end', category)

    def on_income_type_change(self, event):
        """收支类型改变"""
        self.search_var.set('')  # 清空搜索框
        self.update_target_listbox()

    def on_source_category_select(self, event):
        """源分类选择事件"""
        selection = self.source_tree.selection()
        if not selection:
            return

        key = selection[0]
        if key in self.category_mapping:
            if self.category_mapping[key] == '__DELETE__':
                self.current_mapping_label.config(
                    text=f"→ [删除]",
                    foreground='red'
                )
            else:
                self.current_mapping_label.config(
                    text=f"→ {self.category_mapping[key]}",
                    foreground='green'
                )
        else:
            self.current_mapping_label.config(text='未映射', foreground='gray')

    def on_search_category(self, event):
        """搜索目标分类"""
        search_text = self.search_var.get().lower().strip()
        self.target_listbox.delete('0', 'end')

        income_type = self.income_type_var.get()
        if income_type in self.target_categories_data:
            for category in self.target_categories_data[income_type]:
                if search_text == '' or search_text in category.lower():
                    self.target_listbox.insert('end', category)

    def on_source_account_select(self, event):
        """源账户选择事件"""
        selection = self.source_account_listbox.curselection()
        if not selection:
            return

        account = self.source_account_listbox.get(selection[0])
        if account in self.account_mapping:
            self.current_account_mapping_label.config(
                text=f"→ {self.account_mapping[account]}",
                foreground='green'
            )
            self.target_account_entry.delete('0', tk.END)
            self.target_account_entry.insert('0', self.account_mapping[account])
        else:
            self.current_account_mapping_label.config(text='未映射', foreground='gray')
            self.target_account_entry.delete('0', tk.END)

    def set_target_account(self, account):
        """设置目标账户"""
        self.target_account_entry.delete('0', tk.END)
        self.target_account_entry.insert('0', account)

    def set_category_mapping(self):
        """设置分类映射"""
        selection = self.source_tree.selection()
        if not selection:
            messagebox.showwarning('警告', '请先选择源分类')
            return

        target_selection = self.target_listbox.curselection()
        if not target_selection:
            messagebox.showwarning('警告', '请先选择目标分类')
            return

        key = selection[0]
        target_category = self.target_listbox.get(target_selection[0])

        self.category_mapping[key] = target_category
        self.save_mappings()

        # 更新Treeview中的状态
        self.update_source_category_status(key, 'mapped')

        # 更新显示
        self.current_mapping_label.config(
            text=f"→ {target_category}",
            foreground='green'
        )

        messagebox.showinfo('成功', f'已映射: {key} → {target_category}')

    def set_category_delete(self):
        """标记分类为删除"""
        selection = self.source_tree.selection()
        if not selection:
            messagebox.showwarning('警告', '请先选择源分类')
            return

        key = selection[0]
        self.category_mapping[key] = '__DELETE__'
        self.save_mappings()

        # 更新Treeview中的状态
        self.update_source_category_status(key, 'deleted')

        # 更新显示
        self.current_mapping_label.config(
            text=f"→ [删除]",
            foreground='red'
        )

        messagebox.showinfo('成功', f'已标记为删除: {key}')

    def update_source_category_status(self, key, status_type):
        """更新源分类的状态显示"""
        if status_type == 'deleted':
            new_values = list(self.source_tree.item(key)['values'])
            new_values[3] = '已删除'
            self.source_tree.item(key, values=new_values, tags=('deleted',))
        elif status_type == 'mapped':
            new_values = list(self.source_tree.item(key)['values'])
            new_values[3] = '已映射'
            self.source_tree.item(key, values=new_values, tags=('mapped',))
        elif status_type == 'unmapped':
            new_values = list(self.source_tree.item(key)['values'])
            new_values[3] = '未映射'
            self.source_tree.item(key, values=new_values, tags=('unmapped',))

    def clear_category_mapping(self):
        """清除分类映射"""
        selection = self.source_tree.selection()
        if not selection:
            return

        key = selection[0]
        if key in self.category_mapping:
            del self.category_mapping[key]
            self.save_mappings()

        # 更新Treeview中的状态
        self.update_source_category_status(key, 'unmapped')

        self.current_mapping_label.config(text='未映射', foreground='gray')

    def set_account_mapping(self):
        """设置账户映射"""
        selection = self.source_account_listbox.curselection()
        if not selection:
            messagebox.showwarning('警告', '请先选择源账户')
            return

        target_account = self.target_account_entry.get().strip()
        if not target_account:
            messagebox.showwarning('警告', '请输入目标账户名称')
            return

        source_account = self.source_account_listbox.get(selection[0])
        self.account_mapping[source_account] = target_account
        self.save_mappings()

        self.current_account_mapping_label.config(
            text=f"→ {target_account}",
            foreground='green'
        )

        messagebox.showinfo('成功', f'已映射: {source_account} → {target_account}')

    def clear_account_mapping(self):
        """清除账户映射"""
        selection = self.source_account_listbox.curselection()
        if not selection:
            return

        source_account = self.source_account_listbox.get(selection[0])
        if source_account in self.account_mapping:
            del self.account_mapping[source_account]
            self.save_mappings()

        self.current_account_mapping_label.config(text='未映射', foreground='gray')

    def auto_match_categories(self):
        """自动匹配分类"""
        if self.source_data is None or not hasattr(self, 'target_categories_data'):
            messagebox.showwarning('警告', '请先加载源文件和目标分类')
            return

        match_count = 0
        for _, row in self.source_data.iterrows():
            income_type = row['income_type']
            category = row['category']
            sub_category = row['sub_category']

            key = f"{income_type}|{category}|{sub_category}"

            if key not in self.category_mapping:
                # 尝试精确匹配
                if income_type in self.target_categories_data:
                    if category in self.target_categories_data[income_type]:
                        self.category_mapping[key] = category
                        match_count += 1
                    elif sub_category in self.target_categories_data[income_type]:
                        self.category_mapping[key] = sub_category
                        match_count += 1

        self.save_mappings()
        messagebox.showinfo('完成', f'自动匹配完成！匹配了 {match_count} 个分类')

    def preview_conversion(self):
        """预览转换结果"""
        if self.source_data is None:
            messagebox.showwarning('警告', '请先加载源文件')
            return

        self.preview_text.delete('1.0', tk.END)

        # 转换数据
        converted_data = self.convert_data()

        if converted_data is not None and len(converted_data) > 0:
            preview = f"转换后的记录数: {len(converted_data)}\n\n"
            preview += "前10条记录预览:\n"
            preview += "="*100 + "\n"

            for idx, row in converted_data.head(10).iterrows():
                preview += f"{row['交易日期']} | {row['收支类型']} | {row['收支项目']} | {row['金额']} | {row['账户名称']}\n"

            self.preview_text.insert('1.0', preview)
        else:
            self.preview_text.insert('1.0', "没有数据可转换")

    def convert_data(self):
        """转换数据"""
        try:
            result = []
            deleted_count = 0

            for _, row in self.source_data.iterrows():
                # 创建映射键
                key = f"{row['income_type']}|{row['category']}|{row['sub_category']}"

                # 获取目标分类
                if key in self.category_mapping:
                    target_category = self.category_mapping[key]
                    # 检查是否标记为删除
                    if target_category == '__DELETE__':
                        deleted_count += 1
                        continue  # 跳过标记为删除的记录
                elif self.export_unmapped_var.get():
                    target_category = row['category']  # 使用原分类
                else:
                    continue  # 跳过未映射的

                # 获取目标账户
                source_account = row.get('account_name', '')
                if pd.notna(source_account) and source_account in self.account_mapping:
                    target_account = self.account_mapping[source_account]
                else:
                    target_account = source_account if pd.notna(source_account) else '银行卡'

                # 构建新记录
                new_record = {
                    '交易日期': row.get('date', ''),
                    '收支类型': row['income_type'],
                    '收支项目': target_category,
                    '金额': abs(float(row.get('income_amount', 0) if pd.notna(row.get('income_amount')) else 0)),
                    '账户名称': target_account,
                    '标签': '',
                    '备注': row.get('remark', '')
                }

                result.append(new_record)

            print(f"转换完成：共 {len(result)} 条记录，已过滤 {deleted_count} 条删除标记的记录")
            return pd.DataFrame(result)

        except Exception as e:
            messagebox.showerror('错误', f'转换失败: {str(e)}')
            return None

    def export_csv(self):
        """导出CSV"""
        if self.source_data is None:
            messagebox.showwarning('警告', '请先加载源文件')
            return

        converted_data = self.convert_data()

        if converted_data is None or len(converted_data) == 0:
            messagebox.showwarning('警告', '没有数据可导出')
            return

        # 选择保存路径
        file_path = filedialog.asksaveasfilename(
            title='保存转换后的文件',
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )

        if file_path:
            try:
                # 读取模板获取列顺序
                template_columns = ['交易日期', '收支类型', '收支项目', '金额', '账户名称', '标签', '备注']

                # 确保所有列都存在
                for col in template_columns:
                    if col not in converted_data.columns:
                        converted_data[col] = ''

                # 按模板列顺序导出
                converted_data = converted_data[template_columns]
                converted_data.to_csv(file_path, index=False, encoding='utf-8-sig')

                messagebox.showinfo('成功', f'已导出 {len(converted_data)} 条记录到:\n{file_path}')
            except Exception as e:
                messagebox.showerror('错误', f'导出失败: {str(e)}')

def main():
    root = tk.Tk()
    app = PixiuConverterGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()
