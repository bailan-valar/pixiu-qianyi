import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from pathlib import Path
from datetime import datetime
import subprocess
import platform

class PixiuConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("貔貅记账转换工具")
        self.root.geometry("1200x800")

        self.source_file = None
        self.source_data = None
        self.category_mapping = {}
        self.account_mapping = {}
        self.target_accounts = []

        # 加载已保存的映射配置
        self.load_mappings()
        self.load_target_accounts()

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
        ttk.Button(btn_frame, text='删除目标分类', command=self.delete_target_category).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='新增目标分类', command=self.add_new_target_category).pack(side='right', padx=5)

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
        ttk.Button(control_frame, text='导入账户', command=self.import_target_accounts).pack(side='right', padx=5)

        # 分割窗口
        paned = ttk.PanedWindow(frame, orient='horizontal')
        paned.pack(fill='both', expand=True)

        # 左侧：源账户列表
        left_frame = ttk.LabelFrame(paned, text='源账户')
        paned.add(left_frame, weight=1)

        # 创建Treeview显示源账户（添加状态列）
        account_columns = ('account_name', 'status')
        self.source_account_tree = ttk.Treeview(left_frame, columns=account_columns, show='headings', height=25)
        self.source_account_tree.heading('account_name', text='账户名称')
        self.source_account_tree.heading('status', text='状态')

        self.source_account_tree.column('account_name', width=200)
        self.source_account_tree.column('status', width=80)

        # 定义状态标签颜色
        self.source_account_tree.tag_configure('mapped', foreground='green')
        self.source_account_tree.tag_configure('unmapped', foreground='gray')

        source_account_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.source_account_tree.yview)
        self.source_account_tree.configure(yscrollcommand=source_account_scrollbar.set)

        self.source_account_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        source_account_scrollbar.pack(side='right', fill='y', pady=5)

        self.source_account_tree.bind('<<TreeviewSelect>>', self.on_source_account_select)

        # 右侧：目标账户输入
        right_frame = ttk.LabelFrame(paned, text='目标账户设置')
        paned.add(right_frame, weight=1)

        # 账户映射设置
        setting_frame = ttk.Frame(right_frame)
        setting_frame.pack(fill='x', pady=10)

        ttk.Label(setting_frame, text='目标账户名称:').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.target_account_entry = ttk.Entry(setting_frame, width=30)
        self.target_account_entry.grid(row=0, column=1, padx=5, pady=5)

        # 快速选择 - 表格形式
        self.quick_account_frame = ttk.LabelFrame(right_frame, text='选择目标账户')
        self.quick_account_frame.pack(fill='both', expand=True, pady=10, padx=5)

        # 创建Treeview显示账户列表
        account_columns = ('account_type', 'account_owner', 'account_name')
        self.target_account_tree = ttk.Treeview(
            self.quick_account_frame,
            columns=account_columns,
            show='headings',
            height=15
        )

        self.target_account_tree.heading('account_type', text='账户类型')
        self.target_account_tree.heading('account_owner', text='账户所有者')
        self.target_account_tree.heading('account_name', text='账户名称')

        self.target_account_tree.column('account_type', width=100)
        self.target_account_tree.column('account_owner', width=100)
        self.target_account_tree.column('account_name', width=200)

        # 添加滚动条
        account_scrollbar = ttk.Scrollbar(self.quick_account_frame, orient='vertical', command=self.target_account_tree.yview)
        self.target_account_tree.configure(yscrollcommand=account_scrollbar.set)

        self.target_account_tree.pack(side='left', fill='both', expand=True, padx=(5, 0), pady=5)
        account_scrollbar.pack(side='right', fill='y', pady=5)

        # 绑定选择事件
        self.target_account_tree.bind('<<TreeviewSelect>>', self.on_target_account_select)

        # 加载账户数据
        self.load_target_accounts_table()

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
        ttk.Button(btn_frame, text='新增账户', command=self.add_new_account).pack(side='right', padx=5)

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

            # 过滤只保留收入和支出的记录
            original_count = len(self.source_data)
            self.source_data = self.source_data[self.source_data['income_type'].isin(['收入', '支出'])]
            filtered_count = len(self.source_data)

            # 显示文件信息
            info = f"文件: {self.source_file}\n"
            info += f"原始记录数: {original_count}\n"
            info += f"过滤后记录数: {filtered_count} (只保留收入和支出类型)\n"
            info += f"过滤掉记录数: {original_count - filtered_count}\n"
            info += f"列数: {len(self.source_data.columns)}\n\n"
            info += "列名:\n"
            for col in self.source_data.columns:
                info += f"  - {col}\n"

            self.file_info_text.delete('1.0', tk.END)
            self.file_info_text.insert('1.0', info)

            # 加载源分类到Treeview
            self.load_source_categories()
            self.load_source_accounts()

            messagebox.showinfo('成功', f'文件加载成功！\n已过滤掉 {original_count - filtered_count} 条非收入/支出记录')

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
        # 清空Treeview
        for item in self.source_account_tree.get_children():
            self.source_account_tree.delete(item)

        if self.source_data is None:
            return

        # 获取唯一账户 - 根据收支类型选择对应的账户列
        accounts = set()

        for _, row in self.source_data.iterrows():
            if row['income_type'] == '收入':
                # 收入：in_account_name 是我的账户
                account = row.get('in_account_name', '')
            else:  # 支出
                # 支出：from_account_name 是我的账户
                account = row.get('from_account_name', '')

            if pd.notna(account) and account:
                accounts.add(account)

        # 排序并显示
        for account in sorted(accounts):
            # 检查映射状态
            if account in self.account_mapping:
                status = '已映射'
                tag = 'mapped'
            else:
                status = '未映射'
                tag = 'unmapped'

            # 插入数据，包含状态列
            self.source_account_tree.insert('', 'end', iid=account, values=(account, status), tags=(tag,))

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
        selection = self.source_account_tree.selection()
        if not selection:
            return

        account = selection[0]
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
        selection = self.source_account_tree.selection()
        if not selection:
            messagebox.showwarning('警告', '请先选择源账户')
            return

        target_account = self.target_account_entry.get().strip()
        if not target_account:
            messagebox.showwarning('警告', '请输入目标账户名称')
            return

        source_account = selection[0]
        self.account_mapping[source_account] = target_account
        self.save_mappings()

        # 更新Treeview中的状态
        self.update_source_account_status(source_account, 'mapped')

        self.current_account_mapping_label.config(
            text=f"→ {target_account}",
            foreground='green'
        )

    def clear_account_mapping(self):
        """清除账户映射"""
        selection = self.source_account_tree.selection()
        if not selection:
            return

        source_account = selection[0]
        if source_account in self.account_mapping:
            del self.account_mapping[source_account]
            self.save_mappings()

            # 更新Treeview中的状态
            self.update_source_account_status(source_account, 'unmapped')

        self.current_account_mapping_label.config(text='未映射', foreground='gray')

    def update_source_account_status(self, account, status_type):
        """更新源账户的状态显示"""
        if status_type == 'mapped':
            new_values = list(self.source_account_tree.item(account)['values'])
            new_values[1] = '已映射'
            self.source_account_tree.item(account, values=new_values, tags=('mapped',))
        elif status_type == 'unmapped':
            new_values = list(self.source_account_tree.item(account)['values'])
            new_values[1] = '未映射'
            self.source_account_tree.item(account, values=new_values, tags=('unmapped',))

    def load_target_accounts(self):
        """加载目标账户列表"""
        # 尝试从保存的文件加载
        data_needs_migration = False  # 标记是否需要迁移旧格式

        if os.path.exists('target_accounts.json'):
            try:
                with open('target_accounts.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # 数据清理和格式转换
                self.target_accounts_data = []
                if isinstance(data, list):
                    for item in data:
                        if isinstance(item, dict):
                            # 新格式：确保所有必需字段都存在
                            account = {
                                'account_name': item.get('account_name', ''),
                                'account_type': item.get('account_type', '未分类'),
                                'account_owner': item.get('account_owner', '无')
                            }
                            if account['account_name']:  # 只有账户名称不为空才添加
                                self.target_accounts_data.append(account)
                        elif isinstance(item, str):
                            # 旧格式：字符串，转换为字典
                            if item.strip():  # 只有字符串不为空才转换
                                self.target_accounts_data.append({
                                    'account_name': item.strip(),
                                    'account_type': '未分类',
                                    'account_owner': '无'
                                })
                                data_needs_migration = True  # 标记需要迁移
                else:
                    # 不是列表格式，初始化为空
                    self.target_accounts_data = []

                # 如果清理后数据为空，初始化默认数据
                if not self.target_accounts_data:
                    self.target_accounts_data = []

                # 如果检测到旧格式，自动保存为新格式
                if data_needs_migration and self.target_accounts_data:
                    try:
                        with open('target_accounts.json', 'w', encoding='utf-8') as f:
                            json.dump(self.target_accounts_data, f, ensure_ascii=False, indent=2)
                        print("已自动将旧格式账户数据迁移为新格式")
                    except:
                        pass  # 保存失败不影响使用

            except Exception as e:
                print(f"加载目标账户失败: {e}")
                self.target_accounts_data = []
        else:
            self.target_accounts_data = []

    def load_target_accounts_table(self):
        """加载目标账户到表格"""
        # 清空表格
        for item in self.target_account_tree.get_children():
            self.target_account_tree.delete(item)

        # 确保数据格式正确且为列表
        if not isinstance(self.target_accounts_data, list):
            print(f"警告：target_accounts_data 不是列表格式，类型为 {type(self.target_accounts_data)}")
            self.target_accounts_data = []

        # 清理数据：只保留字典格式的有效数据
        cleaned_accounts = []
        for item in self.target_accounts_data:
            if isinstance(item, dict) and item.get('account_name'):
                # 确保所有必需字段都存在
                cleaned_accounts.append({
                    'account_name': str(item.get('account_name', '')),
                    'account_type': str(item.get('account_type', '未分类')),
                    'account_owner': str(item.get('account_owner', '无'))
                })
            elif isinstance(item, str):
                # 处理字符串格式的旧数据
                cleaned_accounts.append({
                    'account_name': item.strip(),
                    'account_type': '未分类',
                    'account_owner': '无'
                })

        # 更新清理后的数据
        self.target_accounts_data = cleaned_accounts

        # 按账户类型和账户名称排序
        try:
            sorted_accounts = sorted(cleaned_accounts, key=lambda x: (x['account_type'], x['account_name']))
        except Exception as e:
            print(f"排序账户失败: {e}")
            sorted_accounts = cleaned_accounts

        # 插入数据到表格
        for account in sorted_accounts:
            account_type = account['account_type']
            account_owner = account['account_owner']
            account_name = account['account_name']

            self.target_account_tree.insert('', 'end', values=(account_type, account_owner, account_name))

    def on_target_account_select(self, event):
        """目标账户表格选择事件"""
        selection = self.target_account_tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.target_account_tree.item(item)['values']
        if values:
            account_name = values[2]  # 账户名称在第3列
            self.set_target_account(account_name)

    def import_target_accounts(self):
        """导入貔貅记账账户文件"""
        file_path = filedialog.askopenfilename(
            title='选择貔貅记账账户文件 (貔貅记账#所有账户信息.csv)',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )

        if not file_path:
            return

        try:
            # 读取CSV文件
            df = pd.read_csv(file_path)

            # 检查必要的列是否存在
            required_columns = ['账户名称', '账户类型', '账户所有者']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                messagebox.showerror('错误', f'CSV文件缺少必要的列: {", ".join(missing_columns)}\n请确保文件包含"账户名称"、"账户类型"、"账户所有者"列。')
                return

            # 提取账户信息
            accounts_data = []
            for _, row in df.iterrows():
                account_name = str(row['账户名称']).strip()
                account_type = str(row['账户类型']).strip()
                account_owner = str(row['账户所有者']).strip()

                if account_name and account_name != 'nan':
                    accounts_data.append({
                        'account_name': account_name,
                        'account_type': account_type if account_type and account_type != 'nan' else '未分类',
                        'account_owner': account_owner if account_owner and account_owner != 'nan' else '无'
                    })

            if not accounts_data:
                messagebox.showwarning('警告', '文件中没有找到有效的账户信息')
                return

            # 更新目标账户列表
            self.target_accounts_data = accounts_data

            # 保存到文件
            try:
                with open('target_accounts.json', 'w', encoding='utf-8') as f:
                    json.dump(self.target_accounts_data, f, ensure_ascii=False, indent=2)
            except:
                pass  # 保存失败不影响继续

            # 更新界面
            self.load_target_accounts_table()

            messagebox.showinfo('成功', f'已导入 {len(accounts_data)} 个账户')

        except Exception as e:
            messagebox.showerror('错误', f'导入账户文件失败: {str(e)}')

    def add_new_target_category(self):
        """新增目标分类"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title('新增目标分类')
        dialog.geometry('400x200')
        dialog.transient(self.root)
        dialog.grab_set()

        # 收支类型选择
        type_frame = ttk.Frame(dialog)
        type_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(type_frame, text='收支类型:').pack(side='left', padx=5)
        new_type_var = tk.StringVar(value=self.income_type_var.get())
        type_combo = ttk.Combobox(type_frame, textvariable=new_type_var, values=['收入', '支出'], state='readonly', width=10)
        type_combo.pack(side='left', padx=5)

        # 新分类名称输入
        name_frame = ttk.Frame(dialog)
        name_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(name_frame, text='分类名称:').pack(side='left', padx=5)
        new_category_entry = ttk.Entry(name_frame, width=30)
        new_category_entry.pack(side='left', padx=5)

        # 检查是否有选中的源分类，如果有则填充默认值
        default_category = ''
        selection = self.source_tree.selection()
        if selection:
            key = selection[0]
            values = self.source_tree.item(key)['values']
            if values:
                # values = (收入类型, 分类, 子分类, 状态)
                category = values[1] if len(values) > 1 else ''
                sub_category = values[2] if len(values) > 2 else ''
                # 如果子分类不是nan，使用子分类；否则使用分类
                if sub_category and sub_category != 'nan':
                    default_category = sub_category
                elif category:
                    default_category = category

        # 填充默认值并选中
        if default_category:
            new_category_entry.insert(0, default_category)
            new_category_entry.select_range(0, 'end')

        new_category_entry.focus()

        # 提示信息
        hint_label = ttk.Label(dialog, text='提示：新增的分类将保存到目标分类列表中', foreground='gray')
        hint_label.pack(pady=5)

        def confirm_add():
            """确认添加"""
            new_category = new_category_entry.get().strip()
            income_type = new_type_var.get()

            if not new_category:
                messagebox.showwarning('警告', '请输入分类名称', parent=dialog)
                return

            # 检查分类是否已存在
            if income_type in self.target_categories_data:
                if new_category in self.target_categories_data[income_type]:
                    messagebox.showwarning('警告', f'分类 "{new_category}" 已存在', parent=dialog)
                    return

            # 添加到目标分类数据
            if income_type not in self.target_categories_data:
                self.target_categories_data[income_type] = []

            self.target_categories_data[income_type].append(new_category)

            # 排序
            self.target_categories_data[income_type].sort()

            # 保存到文件
            try:
                with open('target_category.json', 'w', encoding='utf-8') as f:
                    json.dump(self.target_categories_data, f, ensure_ascii=False, indent=2)

                # 更新界面
                self.income_type_var.set(income_type)
                self.update_target_listbox()

                # 选中新添加的分类
                for i in range(self.target_listbox.size()):
                    if self.target_listbox.get(i) == new_category:
                        self.target_listbox.selection_set(i)
                        self.target_listbox.see(i)
                        break

                # 检查是否有选中的源分类，如果有则自动建立映射
                source_selection = self.source_tree.selection()
                if source_selection:
                    source_key = source_selection[0]
                    self.category_mapping[source_key] = new_category
                    self.save_mappings()
                    # 更新源分类状态显示
                    self.update_source_category_status(source_key, 'mapped')
                    # 更新当前映射标签
                    self.current_mapping_label.config(
                        text=f"→ {new_category}",
                        foreground='green'
                    )

                dialog.destroy()

            except Exception as e:
                messagebox.showerror('错误', f'保存失败: {str(e)}', parent=dialog)

        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text='确认添加', command=confirm_add).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='取消', command=dialog.destroy).pack(side='left', padx=5)

        # 绑定回车键
        new_category_entry.bind('<Return>', lambda e: confirm_add())

    def add_new_account(self):
        """新增账户"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title('新增账户')
        dialog.geometry('500x300')
        dialog.transient(self.root)
        dialog.grab_set()

        # 账户名称输入
        name_frame = ttk.Frame(dialog)
        name_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(name_frame, text='账户名称:', width=12).pack(side='left', padx=5)
        new_account_entry = ttk.Entry(name_frame, width=30)
        new_account_entry.pack(side='left', padx=5)

        # 账户类型选择
        type_frame = ttk.Frame(dialog)
        type_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(type_frame, text='账户类型:', width=12).pack(side='left', padx=5)
        account_types = ['现金', '银行卡', '信用卡', '支付宝', '微信零钱', '虚拟账户', '投资账户', '负债账户', '其他']
        new_account_type_var = tk.StringVar(value='银行卡')
        type_combo = ttk.Combobox(type_frame, textvariable=new_account_type_var, values=account_types, state='readonly', width=27)
        type_combo.pack(side='left', padx=5)

        # 账户所有者输入
        owner_frame = ttk.Frame(dialog)
        owner_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(owner_frame, text='账户所有者:', width=12).pack(side='left', padx=5)
        new_account_owner_var = tk.StringVar(value='无')
        owner_entry = ttk.Entry(owner_frame, textvariable=new_account_owner_var, width=30)
        owner_entry.pack(side='left', padx=5)

        # 检查是否有选中的源账户，如果有则填充默认值
        selection = self.source_account_tree.selection()
        if selection:
            default_account = selection[0]
            new_account_entry.insert(0, default_account)
            new_account_entry.select_range(0, 'end')

        new_account_entry.focus()

        # 提示信息
        hint_label = ttk.Label(dialog, text='提示：新增的账户将保存到目标账户列表中，并自动建立映射', foreground='gray')
        hint_label.pack(pady=10)

        def confirm_add():
            """确认添加"""
            new_account_name = new_account_entry.get().strip()
            account_type = new_account_type_var.get()
            account_owner = new_account_owner_var.get().strip()

            if not new_account_name:
                messagebox.showwarning('警告', '请输入账户名称', parent=dialog)
                return

            # 检查账户是否已存在
            for acc in self.target_accounts_data:
                if acc['account_name'] == new_account_name:
                    messagebox.showwarning('警告', f'账户 "{new_account_name}" 已存在', parent=dialog)
                    return

            # 添加新账户
            new_account = {
                'account_name': new_account_name,
                'account_type': account_type,
                'account_owner': account_owner if account_owner else '无'
            }
            self.target_accounts_data.append(new_account)

            # 保存到文件
            try:
                with open('target_accounts.json', 'w', encoding='utf-8') as f:
                    json.dump(self.target_accounts_data, f, ensure_ascii=False, indent=2)

                # 更新目标账户表格
                self.load_target_accounts_table()

                # 如果有选中的源账户，自动建立映射
                if selection:
                    source_account = selection[0]
                    self.account_mapping[source_account] = new_account_name
                    self.save_mappings()

                    # 更新源账户状态显示
                    self.update_source_account_status(source_account, 'mapped')

                    # 更新当前映射标签
                    self.current_account_mapping_label.config(
                        text=f"→ {new_account_name}",
                        foreground='green'
                    )

                    # 更新目标账户输入框
                    self.target_account_entry.delete('0', tk.END)
                    self.target_account_entry.insert('0', new_account_name)

                dialog.destroy()
                messagebox.showinfo('成功', f'账户 "{new_account_name}" 已添加')

            except Exception as e:
                messagebox.showerror('错误', f'保存失败: {str(e)}', parent=dialog)

        def confirm_and_add_another():
            """确认添加并继续添加"""
            confirm_add()
            if dialog.winfo_exists():
                # 清空输入框，准备添加下一个
                new_account_entry.delete('0', tk.END)
                new_account_entry.focus()
                # 获取当前源账户作为默认值
                selection = self.source_account_tree.selection()
                if selection:
                    new_account_entry.insert(0, selection[0])
                    new_account_entry.select_range(0, 'end')

        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text='确认添加', command=confirm_add).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='添加并继续', command=confirm_and_add_another).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='取消', command=dialog.destroy).pack(side='left', padx=5)

        # 绑定回车键
        new_account_entry.bind('<Return>', lambda e: confirm_add())

    def delete_target_category(self):
        """删除目标分类"""
        selection = self.target_listbox.curselection()
        if not selection:
            messagebox.showwarning('警告', '请先选择要删除的目标分类')
            return

        category_to_delete = self.target_listbox.get(selection[0])
        income_type = self.income_type_var.get()

        # 确认删除
        confirm = messagebox.askyesno(
            '确认删除',
            f'确定要删除目标分类 "{category_to_delete}" 吗？\n\n'
            f'注意：这不会影响已保存的映射配置，只是从目标分类列表中移除。',
            icon='warning'
        )

        if not confirm:
            return

        # 从目标分类数据中删除
        if income_type in self.target_categories_data:
            if category_to_delete in self.target_categories_data[income_type]:
                self.target_categories_data[income_type].remove(category_to_delete)

                # 保存到文件
                try:
                    with open('target_category.json', 'w', encoding='utf-8') as f:
                        json.dump(self.target_categories_data, f, ensure_ascii=False, indent=2)

                    # 更新界面
                    self.update_target_listbox()

                except Exception as e:
                    messagebox.showerror('错误', f'保存失败: {str(e)}')

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

                # 获取目标账户和标签 - 根据收支类型选择不同的账户列
                if row['income_type'] == '收入':
                    source_account = row.get('in_account_name', '')
                    tag_account = row.get('from_account_name', '')
                else:  # 支出
                    source_account = row.get('from_account_name', '')
                    tag_account = row.get('in_account_name', '')

                if pd.notna(source_account) and source_account in self.account_mapping:
                    target_account = self.account_mapping[source_account]
                else:
                    target_account = source_account if pd.notna(source_account) else '银行卡'

                # 构建标签，如果有对方账户则加入到标签中
                tags = []
                if pd.notna(tag_account) and tag_account:
                    tags.append(tag_account)
                original_label = row.get('label', '')
                if pd.notna(original_label) and original_label:
                    tags.append(original_label)

                # 构建新记录
                # 格式化日期：从"2026-02-13 10:46"格式转换为"2026-02-13"
                date_str = str(row.get('date', ''))
                if ' ' in date_str:
                    date_str = date_str.split(' ')[0]  # 只取日期部分

                new_record = {
                    '交易日期': date_str,
                    '收支类型': row['income_type'],
                    '收支项目': target_category,
                    '金额': abs(float(row.get('amount', 0) if pd.notna(row.get('amount')) else 0)),
                    '账户名称': target_account,
                    '标签': ','.join(tags) if tags else '',
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

        # 生成默认文件名：pixiuimport + 时间戳
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f'pixiuimport_{timestamp}.csv'

        # 选择保存路径
        file_path = filedialog.asksaveasfilename(
            title='保存转换后的文件',
            initialfile=default_filename,
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

                # 显示成功对话框
                self.show_export_success_dialog(file_path, len(converted_data))
            except Exception as e:
                messagebox.showerror('错误', f'导出失败: {str(e)}')

    def show_export_success_dialog(self, file_path, record_count):
        """显示导出成功对话框"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title('导出成功')
        dialog.geometry('500x200')
        dialog.transient(self.root)
        dialog.grab_set()

        # 成功图标和消息
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # 成功消息
        success_label = ttk.Label(main_frame, text='✓ 导出成功！', font=('Arial', 16, 'bold'), foreground='green')
        success_label.pack(pady=(10, 20))

        # 文件信息
        info_frame = ttk.LabelFrame(main_frame, text='导出信息')
        info_frame.pack(fill='x', pady=(0, 20))

        ttk.Label(info_frame, text=f'记录数: {record_count} 条').pack(anchor='w', padx=10, pady=5)
        ttk.Label(info_frame, text=f'文件路径:', anchor='w').pack(anchor='w', padx=10, pady=(5, 0))

        # 文件路径（可能很长，使用可滚动的文本框）
        path_text = tk.Text(info_frame, height=3, wrap='word', font=('Arial', 9))
        path_scrollbar = ttk.Scrollbar(info_frame, orient='vertical', command=path_text.yview)
        path_text.configure(yscrollcommand=path_scrollbar.set)

        path_text.insert('1.0', file_path)
        path_text.config(state='disabled')  # 只读
        path_text.pack(fill='x', padx=10, pady=(0, 10))
        path_scrollbar.pack(side='right', fill='y')

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(0, 10))

        def open_file():
            """打开文件"""
            try:
                if platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', file_path])
                elif platform.system() == 'Windows':
                    os.startfile(file_path)
                else:  # Linux
                    subprocess.run(['xdg-open', file_path])
            except Exception as e:
                messagebox.showerror('错误', f'无法打开文件: {str(e)}', parent=dialog)

        def open_folder():
            """打开所在文件夹"""
            try:
                folder_path = os.path.dirname(file_path)
                if platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', folder_path])
                elif platform.system() == 'Windows':
                    subprocess.run(['explorer', folder_path])
                else:  # Linux
                    subprocess.run(['xdg-open', folder_path])
            except Exception as e:
                messagebox.showerror('错误', f'无法打开文件夹: {str(e)}', parent=dialog)

        def close_dialog():
            """关闭对话框"""
            dialog.destroy()

        # 添加按钮
        ttk.Button(btn_frame, text='打开文件', command=open_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='打开所在文件夹', command=open_folder).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='关闭', command=close_dialog).pack(side='right', padx=5)

        # 设置焦点到关闭按钮
        dialog.bind('<Escape>', lambda e: close_dialog())

def main():
    root = tk.Tk()
    app = PixiuConverterGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()
