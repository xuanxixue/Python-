import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import docx
import os
import spacy
from collections import OrderedDict

class DocumentProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("文档句子提取工具")
        self.root.geometry("1000x700")
        
        # 存储上传的文件内容
        self.uploaded_files_content = ""
        self.all_sentences = []  # 存储所有提取的句子（不去重）
        
        # 加载spaCy模型（尝试加载中文，失败则加载英文）
        try:
            self.nlp = spacy.load("zh_core_web_sm")
        except OSError:
            try:
                self.nlp = spacy.load("en_core_web_sm")
                messagebox.showwarning("语言模型", "未找到中文模型，已使用英文模型。如需处理中文，请安装中文模型：python -m spacy download zh_core_web_sm")
            except OSError:
                messagebox.showerror("错误", "未找到spaCy语言模型。请安装至少一种语言模型。")
                self.nlp = None
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="文档句子提取工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 上传按钮区域
        upload_frame = ttk.LabelFrame(main_frame, text="文档上传", padding="10")
        upload_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.upload_btn = ttk.Button(upload_frame, text="上传文档 (docx/txt)", command=self.upload_files)
        self.upload_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 文件列表显示区域
        file_list_frame = ttk.Frame(upload_frame)
        file_list_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        file_list_label = ttk.Label(file_list_frame, text="已上传文件列表:")
        file_list_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 创建文件列表框
        self.file_listbox = tk.Listbox(file_list_frame, width=80, height=6)
        self.file_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        file_list_scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical", command=self.file_listbox.yview)
        file_list_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.file_listbox.configure(yscrollcommand=file_list_scrollbar.set)
        
        # 处理按钮
        self.process_btn = ttk.Button(main_frame, text="处理文档", command=self.process_documents, state="disabled")
        self.process_btn.grid(row=2, column=0, columnspan=2, pady=10)
        
        # 导出按钮区域
        export_frame = ttk.LabelFrame(main_frame, text="导出选项", padding="10")
        export_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.export_txt_btn = ttk.Button(export_frame, text="导出为TXT", command=self.export_txt, state="disabled")
        self.export_txt_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="提取结果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 统计信息
        self.stats_label = ttk.Label(result_frame, text="请先上传并处理文档")
        self.stats_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # 创建表格框架
        table_frame = ttk.Frame(result_frame)
        table_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建表格
        self.create_table(table_frame)
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)
        file_list_frame.columnconfigure(0, weight=1)
        file_list_frame.rowconfigure(1, weight=1)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
    
    def create_table(self, parent):
        # 创建表格
        columns = ("序号", "句子")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)
        
        # 设置列标题
        self.tree.heading("序号", text="序号")
        self.tree.heading("句子", text="句子")
        
        # 设置列宽
        self.tree.column("序号", width=60, anchor="center")
        self.tree.column("句子", width=800, anchor="w")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def upload_files(self):
        """上传多个文档文件"""
        file_types = [
            ("Word文档", "*.docx"),
            ("文本文件", "*.txt"),
            ("所有文件", "*.*")
        ]
        
        file_paths = filedialog.askopenfilenames(
            title="选择文档文件",
            filetypes=file_types
        )
        
        if file_paths:
            # 清空现有文件列表
            self.file_listbox.delete(0, tk.END)
            
            # 添加新文件到列表
            for file_path in file_paths:
                self.file_listbox.insert(tk.END, file_path)
            
            self.process_btn.config(state="normal")
            self.file_paths = file_paths
        else:
            self.process_btn.config(state="disabled")
    
    def read_document(self, file_path):
        """读取文档内容"""
        try:
            if file_path.endswith('.docx'):
                # 读取Word文档
                doc = docx.Document(file_path)
                content = []
                for paragraph in doc.paragraphs:
                    if paragraph.text.strip():
                        content.append(paragraph.text)
                return '\n'.join(content)
            elif file_path.endswith('.txt'):
                # 读取文本文件
                with open(file_path, 'r', encoding='utf-8') as file:
                    return file.read()
            else:
                messagebox.showerror("错误", f"不支持的文件格式: {file_path}")
                return ""
        except Exception as e:
            messagebox.showerror("错误", f"读取文件 {file_path} 时出错: {str(e)}")
            return ""
    
    def merge_documents(self):
        """合并所有上传的文档为无段落文档"""
        combined_content = []
        for file_path in self.file_paths:
            content = self.read_document(file_path)
            if content:
                combined_content.append(content)
        
        self.uploaded_files_content = '\n'.join(combined_content)
        return self.uploaded_files_content
    
    def extract_sentences_spacy(self, text):
        """使用spaCy库提取句子"""
        if self.nlp is None:
            messagebox.showerror("错误", "spaCy模型未正确加载")
            return []
        
        # 处理文本
        doc = self.nlp(text)
        
        # 提取句子
        sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
        
        return sentences
    
    def process_documents(self):
        """处理文档并显示结果"""
        if not hasattr(self, 'file_paths') or not self.file_paths:
            messagebox.showwarning("警告", "请先上传文档文件")
            return
        
        try:
            # 合并文档
            self.merge_documents()
            
            if not self.uploaded_files_content:
                messagebox.showwarning("警告", "没有可处理的内容")
                return
            
            # 使用spaCy提取句子（不去重）
            self.all_sentences = self.extract_sentences_spacy(self.uploaded_files_content)
            
            # 启用导出按钮
            self.export_txt_btn.config(state="normal")
            
            # 显示结果
            self.display_results()
            
        except Exception as e:
            messagebox.showerror("错误", f"处理文档时出错: {str(e)}")
    
    def display_results(self):
        """在表格中显示处理结果"""
        # 更新统计信息
        stats_text = f"提取句子数量: {len(self.all_sentences)}"
        self.stats_label.config(text=stats_text)
        
        # 清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加数据到表格
        if self.all_sentences:
            for i, sentence in enumerate(self.all_sentences, 1):
                self.tree.insert("", "end", values=(i, sentence))
        else:
            messagebox.showinfo("提示", "未提取到任何句子")
    
    def export_txt(self):
        """导出结果为TXT文件"""
        if not self.all_sentences:
            messagebox.showwarning("警告", "没有可导出的句子，请先处理文档")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存为TXT文件",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(f"文档句子提取结果\n")
                    file.write(f"原始文档数量: {len(self.file_paths)}\n")
                    file.write(f"提取句子数量: {len(self.all_sentences)}\n")
                    file.write("=" * 50 + "\n\n")
                    
                    for i, sentence in enumerate(self.all_sentences, 1):
                        file.write(f"{i}. {sentence}\n\n")
                
                messagebox.showinfo("成功", f"句子已成功导出到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出TXT文件时出错: {str(e)}")

def main():
    root = tk.Tk()
    app = DocumentProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()