import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import requests
import threading
import time
import tkinter.ttk as ttk  # 正确导入 ttk 模块

# OpenAI API 配置
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"  # 默认 OpenAI API 地址
OPENAI_API_KEY = "your_openai_api_key"  # 替换为您的 OpenAI API 密钥

# 全局变量，用于取消操作
cancel_processing = False

# 当前选中的模型
selected_model = "gpt-3.5-turbo"

# 进度条相关变量
total_files = 0
processed_files = 0


def extract_text_from_pdf(pdf_path):
    """
    从 PDF 文件中提取文本内容
    """
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
    except Exception as e:
        print(f"无法解析 PDF 文件 {pdf_path}，错误：{e}")
    return text


def ask_questions_to_ai(pdf_text, questions, model):
    """
    使用 OpenAI API 回答问题
    """
    answers = []
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    for question in questions:
        if cancel_processing:
            print("处理已取消")
            return []
        prompt = f"根据以下文本回答问题：\n\n{pdf_text}\n\n问题：{question}"
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7
        }
        try:
            # 设置请求超时时间为 30 秒
            response = requests.post(OPENAI_API_URL, json=payload, headers=headers, timeout=30)
            response.raise_for_status()
            answer = response.json()["choices"][0]["message"]["content"].strip()
        except requests.Timeout:
            print(f"请求超时：{question}")
            answer = "请求超时"
        except Exception as e:
            print(f"无法回答问题 {question}，错误：{e}")
            answer = "无法回答"
        answers.append(answer)
    return answers


def save_to_excel(questions, answers, output_file):
    """
    将问题和答案保存到 Excel 文件中
    """
    df = pd.DataFrame({"问题": questions, "答案": answers})
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"结果已保存到 {output_file}")


def process_pdf(pdf_path, questions, output_dir, progress_label, progress_bar):
    """
    处理单个 PDF 文件，并更新进度条和状态信息
    """
    global processed_files
    print(f"正在处理文件：{pdf_path}")
    progress_label.config(text=f"正在处理：{os.path.basename(pdf_path)}")
    pdf_text = extract_text_from_pdf(pdf_path)
    if not pdf_text:
        print(f"文件 {pdf_path} 无内容，跳过处理")
        processed_files += 1
        progress_bar.step(1)
        return
    answers = ask_questions_to_ai(pdf_text, questions, selected_model)
    if cancel_processing:
        return
    output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx")
    save_to_excel(questions, answers, output_file)
    processed_files += 1
    progress_bar.step(1)


def batch_process_pdfs(pdf_dir, output_dir, questions, progress_label, progress_bar):
    """
    批量处理 PDF 文件，每隔一分钟处理一个 PDF 文件
    """
    global total_files, processed_files
    total_files = 0
    processed_files = 0
    pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith(".pdf")]
    total_files = len(pdf_files)
    progress_bar.config(maximum=total_files)
    progress_bar["value"] = 0

    for file_name in pdf_files:
        if cancel_processing:
            break
        pdf_path = os.path.join(pdf_dir, file_name)
        process_pdf(pdf_path, questions, output_dir, progress_label, progress_bar)
        if not cancel_processing:
            print("等待一分钟后处理下一个文件...")
            progress_label.config(text="等待一分钟后处理下一个文件...")
            time.sleep(60)  # 暂停 60 秒
    progress_label.config(text="处理完成！")


class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 处理工具")
        self.root.geometry("600x600")  # 增大窗口高度

        # PDF 文件夹选择
        self.pdf_dir_label = tk.Label(root, text="PDF 文件夹：")
        self.pdf_dir_label.pack()
        self.pdf_dir_entry = tk.Entry(root, width=50)
        self.pdf_dir_entry.pack()
        self.pdf_dir_button = tk.Button(root, text="选择文件夹", command=self.select_pdf_dir)
        self.pdf_dir_button.pack()

        # 输出文件夹选择
        self.output_dir_label = tk.Label(root, text="输出文件夹：")
        self.output_dir_label.pack()
        self.output_dir_entry = tk.Entry(root, width=50)
        self.output_dir_entry.pack()
        self.output_dir_button = tk.Button(root, text="选择文件夹", command=self.select_output_dir)
        self.output_dir_button.pack()

        # 输入问题
        self.questions_label = tk.Label(root, text="输入问题（每行一个问题）：")
        self.questions_label.pack()
        self.questions_text = tk.Text(root, height=10, width=50)
        self.questions_text.pack()

        # API 地址输入
        self.api_url_label = tk.Label(root, text="OpenAI API 地址：")
        self.api_url_label.pack()
        self.api_url_entry = tk.Entry(root, width=50)
        self.api_url_entry.pack()
        self.update_api_url_button = tk.Button(root, text="更新 API 地址", command=self.update_api_url)
        self.update_api_url_button.pack()

        # API 密钥输入
        self.api_key_label = tk.Label(root, text="OpenAI API 密钥：")
        self.api_key_label.pack()
        self.api_key_entry = tk.Entry(root, width=50)
        self.api_key_entry.pack()
        self.update_api_key_button = tk.Button(root, text="更新 API 密钥", command=self.update_api_key)
        self.update_api_key_button.pack()

        # 模型选择
        self.model_label = tk.Label(root, text="选择模型：")
        self.model_label.pack()
        self.model_var = tk.StringVar()
        self.model_dropdown = ttk.Combobox(root, textvariable=self.model_var, values=["gpt-3.5-turbo", "gpt-4"])
        self.model_dropdown.pack()
        self.model_dropdown.set("gpt-3.5-turbo")  # 默认选中
        self.model_dropdown.bind("<<ComboboxSelected>>", self.update_selected_model)

        # 进度条和状态信息
        self.progress_label = tk.Label(root, text="等待处理...")
        self.progress_label.pack()
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack()

        # 开始处理按钮
        self.start_button = tk.Button(root, text="开始处理", command=self.start_processing)
        self.start_button.pack()

        # 取消按钮
        self.cancel_button = tk.Button(root, text="取消", command=self.cancel_processing)
        self.cancel_button.pack()

    def select_pdf_dir(self):
        """选择 PDF 文件夹"""
        folder_path = filedialog.askdirectory()
        self.pdf_dir_entry.delete(0, tk.END)
        self.pdf_dir_entry.insert(0, folder_path)

    def select_output_dir(self):
        """选择输出文件夹"""
        folder_path = filedialog.askdirectory()
        self.output_dir_entry.delete(0, tk.END)
        self.output_dir_entry.insert(0, folder_path)

    def update_selected_model(self, event):
        """更新选中的模型"""
        global selected_model
        selected_model = self.model_var.get()

    def update_api_url(self):
        """更新 OpenAI API 地址"""
        global OPENAI_API_URL
        new_api_url = self.api_url_entry.get()
        if new_api_url:
            OPENAI_API_URL = new_api_url
            messagebox.showinfo("提示", "API 地址已更新！")
        else:
            messagebox.showerror("错误", "API 地址不能为空！")

    def update_api_key(self):
        """更新 OpenAI API 密钥"""
        global OPENAI_API_KEY
        new_api_key = self.api_key_entry.get()
        if new_api_key:
            OPENAI_API_KEY = new_api_key
            messagebox.showinfo("提示", "API 密钥已更新！")
        else:
            messagebox.showerror("错误", "API 密钥不能为空！")

    def start_processing(self):
        """开始处理 PDF 文件"""
        global cancel_processing, OPENAI_API_URL, OPENAI_API_KEY, selected_model
        cancel_processing = False

        pdf_dir = self.pdf_dir_entry.get()
        output_dir = self.output_dir_entry.get()
        questions = self.questions_text.get("1.0", tk.END).strip().split("\n")
        self.update_selected_model(None)  # 获取当前选中的模型

        if not pdf_dir or not output_dir or not questions:
            messagebox.showerror("错误", "请确保选择了 PDF 文件夹、输出文件夹，并输入了问题。")
            return

        if not os.path.exists(pdf_dir) or not os.path.exists(output_dir):
            messagebox.showerror("错误", "PDF 文件夹或输出文件夹不存在。")
            return

        messagebox.showinfo("提示", "处理开始，请稍候...")
        self.progress_label.config(text="正在初始化...")
        processing_thread = threading.Thread(target=batch_process_pdfs, args=(pdf_dir, output_dir, questions, self.progress_label, self.progress_bar))
        processing_thread.start()

    def cancel_processing(self):
        """取消处理操作"""
        global cancel_processing
        cancel_processing = True
        self.progress_label.config(text="处理已取消")
        messagebox.showinfo("提示", "处理已取消")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()