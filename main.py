# 使用之前需要pip install PyMuPDF
# pip install requests
import requests
import os
import glob
import fitz
import html
from urllib import parse
import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

class Doc88Downloader:
    def __init__(self, root):
        self.root = root
        self.root.title("道客巴巴文档下载器")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("微软雅黑", 10))
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", font=("微软雅黑", 10))
        
        # 创建界面
        self.create_widgets()
        
        # 下载状态
        self.downloading = False
        
        # 创建具有重试功能的session
        self.setup_session()
        
    def setup_session(self):
        # 创建自定义session，带有重试机制
        self.session = requests.Session()
        retry_strategy = Retry(
            total=5,  # 最多重试5次
            backoff_factor=1,  # 重试间隔时间系数
            status_forcelist=[429, 500, 502, 503, 504],  # 特定HTTP状态码的重试
            allowed_methods=["GET", "POST"]  # 允许重试的请求方法
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        
        # 设置默认超时时间
        self.timeout = 30
        
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建输入框和按钮
        ttk.Label(main_frame, text="请输入文档网址或p_code:").pack(anchor=tk.W, pady=(0, 5))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.url_entry = ttk.Entry(input_frame, width=50)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        download_btn = ttk.Button(input_frame, text="下载", command=self.start_download)
        download_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 创建保存位置选择
        save_frame = ttk.Frame(main_frame)
        save_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(save_frame, text="保存位置:").pack(side=tk.LEFT)
        
        self.save_path = os.getcwd()
        self.save_path_var = tk.StringVar(value=self.save_path)
        save_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, width=40)
        save_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(save_frame, text="浏览...", command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT)
        
        # 创建进度条和状态标签
        ttk.Label(main_frame, text="下载进度:").pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=560, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(5, 10))
        
        self.status_var = tk.StringVar(value="准备就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, foreground="gray")
        status_label.pack(anchor=tk.W)
        
        # 创建日志文本框
        log_frame = ttk.LabelFrame(main_frame, text="日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=8, width=65, font=("微软雅黑", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.config(state=tk.DISABLED)
        
    def browse_folder(self):
        folder = filedialog.askdirectory(title="选择保存位置")
        if folder:
            self.save_path = folder
            self.save_path_var.set(folder)
    
    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def update_status(self, message, is_error=False):
        self.status_var.set(message)
        self.log(message)
        if is_error:
            messagebox.showerror("错误", message)
    
    def extract_p_code(self, url_or_code):
        # 如果是URL，提取p_code
        if "doc88.com" in url_or_code:
            try:
                # 分析URL
                if "/p-" in url_or_code:
                    p_code = url_or_code.split("/p-")[1].split(".")[0]
                    return p_code
            except:
                return None
        # 否则假设直接是p_code
        return url_or_code.strip()
    
    def start_download(self):
        if self.downloading:
            messagebox.showinfo("提示", "已有下载任务正在进行中")
            return
            
        url_or_code = self.url_entry.get().strip()
        if not url_or_code:
            messagebox.showwarning("警告", "请输入文档网址或p_code")
            return
            
        p_code = self.extract_p_code(url_or_code)
        if not p_code:
            messagebox.showwarning("警告", "无法识别输入的网址或p_code")
            return
            
        # 使用线程避免UI卡死
        self.downloading = True
        download_thread = threading.Thread(target=self.download_document, args=(p_code,))
        download_thread.daemon = True
        download_thread.start()
    
    def download_document(self, p_code):
        try:
            self.update_status(f"开始下载文档，p_code: {p_code}")
            self.progress["value"] = 0
            
            # 第一步：获取返回编码后的data
            header = {
                'Referer': 'http://m.doc88.com/',
                'User-Agent': 'Mozilla/5.0 (Linux; Android 9; ONEPLUS A6010) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.89 Mobile Safari/537.36'
            }
            url = f'https://m.doc88.com/doc.php?act=info&p_code={p_code}&key=3854933de90d1dbb321d8ca29eac130a&v=1'
            
            self.update_status("正在获取文档信息...")
            result = self.session.get(url, headers=header, timeout=self.timeout)
            return_data = result.text
            
            # 第二步：解码data获取gif信息
            self.update_status("正在解析文档结构...")
            base64str = self.decode_base64(return_data)
            s = json.loads(base64str)
            
            gif_host = s['gif_host']
            file_name = s['name']
            gif_urls = json.loads(s['gif_struct'])
            
            # 创建保存目录
            save_dir = os.path.join(self.save_path, file_name)
            if os.path.exists(save_dir):
                self.update_status(f"目录 {file_name} 已存在，将覆盖其中文件")
            else:
                os.makedirs(save_dir, exist_ok=True)
                
            # 第三步：下载gif文件
            self.update_status(f"开始下载图片，共 {len(gif_urls)} 页")
            
            for index, element in enumerate(gif_urls):
                if not self.downloading:  # 检查是否取消下载
                    break
                    
                gif_url = gif_host + "/get-" + element['u'] + ".gif"
                self.update_status(f"正在下载第 {index+1}/{len(gif_urls)} 页")
                
                # 添加重试逻辑
                max_retries = 3
                retry_count = 0
                success = False
                
                while retry_count < max_retries and not success:
                    try:
                        result = self.session.get(gif_url, timeout=self.timeout, verify=True)
                        result.raise_for_status()  # 检查HTTP状态码
                        
                        with open(os.path.join(save_dir, f'{index:07d}.gif'), 'wb') as f:
                            f.write(result.content)
                        
                        success = True
                    except requests.exceptions.SSLError as ssl_err:
                        retry_count += 1
                        self.log(f"SSL错误，第{retry_count}次尝试重新下载: {str(ssl_err)}")
                        time.sleep(2)  # 等待2秒后重试
                    except requests.exceptions.RequestException as req_err:
                        retry_count += 1
                        self.log(f"请求错误，第{retry_count}次尝试重新下载: {str(req_err)}")
                        time.sleep(2)  # 等待2秒后重试
                
                if not success:
                    self.log(f"下载第{index+1}页失败，跳过此页")
                
                # 更新进度条
                self.progress["value"] = (index + 1) / len(gif_urls) * 80
                self.root.update_idletasks()
            
            # 第四步：合并为PDF
            if self.downloading:
                self.update_status("正在生成PDF...")
                self.pic2pdf(save_dir, file_name)
                
                self.progress["value"] = 100
                self.update_status(f"下载完成！PDF已保存至: {os.path.join(save_dir, file_name)}.pdf")
                messagebox.showinfo("完成", f"文档已下载完成并转换为PDF！\n保存位置: {os.path.join(save_dir, file_name)}.pdf")
        
        except Exception as e:
            self.update_status(f"下载过程中出错: {str(e)}", is_error=True)
        
        finally:
            self.downloading = False
    
    def decode_base64(self, s):
        # 解码函数
        m_base64Str = ''
        m_base64Count = 0
        m_END_OF_INPUT = -1
        m_base64Chars_r = [
            'P', 'J', 'K', 'L', 'M', 'N', 'O', 'I',
            '3', 'y', 'x', 'z', '0', '1', '2', 'w',
            'v', 'p', 'r', 'q', 's', 't', 'u', 'o',
            'B', 'H', 'C', 'D', 'E', 'F', 'G', 'A',
            'h', 'n', 'i', 'j', 'k', 'l', 'm', 'g',
            'f', 'Z', 'a', 'b', 'c', 'd', 'e', 'Y',
            'X', 'R', 'S', 'T', 'U', 'V', 'W', 'Q',
            '!', '5', '6', '7', '8', '9', '+', '4'
        ]
        m_reverseBase64Chars = {element: index for index, element in enumerate(m_base64Chars_r)}
        
        def m_setBase64Str(s):
            nonlocal m_base64Count, m_base64Str
            m_base64Str = s
            m_base64Count = 0
            
        def m_readReverseBase64():
            nonlocal m_base64Count, m_base64Str, m_reverseBase64Chars
            if not m_base64Str:
                return -1
            while 1:
                if m_base64Count >= len(m_base64Str):
                    return -1
                nextCharacter = m_base64Str[m_base64Count]
                m_base64Count += 1
                try:
                    if m_reverseBase64Chars[nextCharacter]:
                        return m_reverseBase64Chars[nextCharacter]
                except:
                    pass
                if nextCharacter == 'P':
                    return 0
            return -1
            
        def m_ntos(n):
            n = hex(n)
            n = n[2:]
            if len(n) == 1:
                n = "0" + n[-1]
            n = '%' + n
            return html.unescape(n)
            
        m_setBase64Str(s)
        result = ''
        done = False
        inBuffer = [0, 0, 0, 0]
        inBuffer[0] = m_readReverseBase64()
        inBuffer[1] = m_readReverseBase64()
        while (not done) and (inBuffer[0] != m_END_OF_INPUT) and (inBuffer[1] != m_END_OF_INPUT):
            inBuffer[2] = m_readReverseBase64()
            inBuffer[3] = m_readReverseBase64()
            result += m_ntos((((inBuffer[0] << 2) & 0xff) | inBuffer[1] >> 4))
            if inBuffer[2] != m_END_OF_INPUT:
                result += m_ntos((((inBuffer[1] << 4) & 0xff) | inBuffer[2] >> 2))
                if inBuffer[3] != m_END_OF_INPUT:
                    result += m_ntos((((inBuffer[2] << 6) & 0xff) | inBuffer[3]))
                else:
                    done = True
            else:
                done = True
            inBuffer[0] = m_readReverseBase64()
            inBuffer[1] = m_readReverseBase64()
        return parse.unquote(result)
    
    def pic2pdf(self, save_dir, file_name):
        doc = fitz.open()
        for img in sorted(glob.glob(os.path.join(save_dir, "*gif"))):
            self.log(f"处理图片: {os.path.basename(img)}")
            imgdoc = fitz.open(img)
            pdfbytes = imgdoc.convert_to_pdf()
            imgpdf = fitz.open("pdf", pdfbytes)
            doc.insert_pdf(imgpdf)
            
        pdf_path = os.path.join(save_dir, f"{file_name}.pdf")
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        doc.save(pdf_path)
        doc.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = Doc88Downloader(root)
    root.mainloop()
