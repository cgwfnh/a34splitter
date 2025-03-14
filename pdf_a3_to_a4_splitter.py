'''
Author: cgwfnh
Version: 0.1
Date: 20250314
Function: 将横版A3大小的PDF文件，分割成2张A4大小的PDF文件，并将分割后的文件合并为单个PDF文件
'''
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk
import io
import tempfile

class PDFA3ToA4Splitter:
    def __init__(self, root):
        self.root = root
        self.root.title("A3 PDF 分割器")
        self.root.geometry("1000x700")
        
        self.pdf_path = None
        self.current_page = 0
        self.total_pages = 0
        self.split_ratio = 0.5  # 默认从中间分割
        self.doc = None
        self.temp_files = []
        self.dragging = False
        self.last_x = 0
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部控制区域
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)
        
        ttk.Button(control_frame, text="选择PDF文件", command=self.select_pdf).grid(row=0, column=0, padx=5)
        ttk.Label(control_frame, text="分割位置:").grid(row=0, column=1, padx=5)
        
        self.split_scale = ttk.Scale(control_frame, from_=0.1, to=0.9, orient=tk.HORIZONTAL, 
                                     length=200, value=0.5, command=self.update_preview)
        self.split_scale.grid(row=0, column=2, padx=5)
        
        self.split_value_label = ttk.Label(control_frame, text="50%")
        self.split_value_label.grid(row=0, column=3, padx=5)
        
        ttk.Button(control_frame, text="上一页", command=self.prev_page).grid(row=0, column=4, padx=5)
        
        self.page_label = ttk.Label(control_frame, text="0/0")
        self.page_label.grid(row=0, column=5, padx=5)
        
        ttk.Button(control_frame, text="下一页", command=self.next_page).grid(row=0, column=6, padx=5)
        
        ttk.Button(control_frame, text="分割并保存", command=self.split_and_save).grid(row=0, column=7, padx=5)
        
        # 预览区域
        preview_frame = ttk.Frame(self.root, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # 原始PDF预览
        original_frame = ttk.LabelFrame(preview_frame, text="原始A3 PDF", padding="10")
        original_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.original_canvas = tk.Canvas(original_frame, bg="white")
        self.original_canvas.pack(fill=tk.BOTH, expand=True)
        
        # 分割预览
        split_frame = ttk.LabelFrame(preview_frame, text="分割预览", padding="10")
        split_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        split_canvas_frame = ttk.Frame(split_frame)
        split_canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.left_canvas = tk.Canvas(split_canvas_frame, bg="white")
        self.left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.right_canvas = tk.Canvas(split_canvas_frame, bg="white")
        self.right_canvas.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar(value="请选择PDF文件")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.pdf_path = file_path
            self.load_pdf()
    
    def load_pdf(self):
        try:
            # 关闭之前打开的文档
            if self.doc:
                self.doc.close()
            
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            
            if self.total_pages > 0:
                self.status_var.set(f"已加载: {os.path.basename(self.pdf_path)}")
                self.page_label.config(text=f"1/{self.total_pages}")
                self.update_preview()
            else:
                self.status_var.set("PDF文件没有页面")
        except Exception as e:
            messagebox.showerror("错误", f"加载PDF时出错: {str(e)}")
            self.status_var.set("加载PDF失败")
    
    def update_preview(self, *args):
        if not self.doc or self.current_page >= self.total_pages:
            return
        
        # 更新分割比例标签
        ratio = self.split_scale.get()
        self.split_ratio = ratio
        self.split_value_label.config(text=f"{int(ratio * 100)}%")
        
        # 获取当前页面
        page = self.doc[self.current_page]
        
        # 清除画布
        self.original_canvas.delete("all")
        self.left_canvas.delete("all")
        self.right_canvas.delete("all")
        
        # 渲染原始页面
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        img_data = pix.tobytes("ppm")
        img = Image.open(io.BytesIO(img_data))
        
        # 调整图像大小以适应画布
        canvas_width = self.original_canvas.winfo_width()
        canvas_height = self.original_canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:  # 画布尚未完全初始化
            self.root.after(100, self.update_preview)
            return
        
        img_ratio = img.width / img.height
        canvas_ratio = canvas_width / canvas_height
        
        if img_ratio > canvas_ratio:
            new_width = canvas_width
            new_height = int(canvas_width / img_ratio)
        else:
            new_height = canvas_height
            new_width = int(canvas_height * img_ratio)
        
        img_resized = img.resize((new_width, new_height), Image.LANCZOS)
        self.original_img = ImageTk.PhotoImage(img_resized)
        
        # 在原始画布上显示图像
        self.original_canvas.create_image(
            canvas_width // 2, canvas_height // 2,
            image=self.original_img, anchor=tk.CENTER
        )
        
        # 显示分割线
        split_x = int(new_width * ratio)
        self.split_line = self.original_canvas.create_line(
            split_x, 0, split_x, new_height,
            fill="red", width=2, dash=(4, 4), tags="split_line"
        )
        
        # 绑定鼠标事件到分割线
        self.original_canvas.tag_bind("split_line", "<ButtonPress-1>", self.start_drag)
        self.original_canvas.tag_bind("split_line", "<B1-Motion>", self.drag_split_line)
        self.original_canvas.tag_bind("split_line", "<ButtonRelease-1>", self.end_drag)
        
        # 创建分割后的图像
        left_img = img.crop((0, 0, int(img.width * ratio), img.height))
        right_img = img.crop((int(img.width * ratio), 0, img.width, img.height))
        
        # 调整分割后图像大小以适应画布
        split_canvas_width = self.left_canvas.winfo_width()
        split_canvas_height = self.left_canvas.winfo_height()
        
        if split_canvas_width <= 1 or split_canvas_height <= 1:
            self.root.after(100, self.update_preview)
            return
        
        left_ratio = left_img.width / left_img.height
        right_ratio = right_img.width / right_img.height
        
        if left_ratio > canvas_ratio:
            left_new_width = split_canvas_width
            left_new_height = int(split_canvas_width / left_ratio)
        else:
            left_new_height = split_canvas_height
            left_new_width = int(split_canvas_height * left_ratio)
        
        if right_ratio > canvas_ratio:
            right_new_width = split_canvas_width
            right_new_height = int(split_canvas_width / right_ratio)
        else:
            right_new_height = split_canvas_height
            right_new_width = int(split_canvas_height * right_ratio)
        
        left_img_resized = left_img.resize((left_new_width, left_new_height), Image.LANCZOS)
        right_img_resized = right_img.resize((right_new_width, right_new_height), Image.LANCZOS)
        
        self.left_img = ImageTk.PhotoImage(left_img_resized)
        self.right_img = ImageTk.PhotoImage(right_img_resized)
        
        # 在分割画布上显示图像
        self.left_canvas.create_image(
            split_canvas_width // 2, split_canvas_height // 2,
            image=self.left_img, anchor=tk.CENTER
        )
        
        self.right_canvas.create_image(
            split_canvas_width // 2, split_canvas_height // 2,
            image=self.right_img, anchor=tk.CENTER
        )
    
    def start_drag(self, event):
        self.dragging = True
        self.last_x = event.x
    
    def drag_split_line(self, event):
        if not self.dragging or not self.doc:
            return
            
        # 获取画布尺寸
        canvas_width = self.original_canvas.winfo_width()
        
        # 计算新的分割位置
        new_x = event.x
        
        # 限制在合理范围内 (10%-90%)
        if new_x < canvas_width * 0.1:
            new_x = canvas_width * 0.1
        elif new_x > canvas_width * 0.9:
            new_x = canvas_width * 0.9
            
        # 移动分割线
        self.original_canvas.coords("split_line", new_x, 0, new_x, self.original_canvas.winfo_height())
        
        # 更新分割比例
        self.split_ratio = new_x / canvas_width
        self.split_scale.set(self.split_ratio)
        self.split_value_label.config(text=f"{int(self.split_ratio * 100)}%")
        
        # 更新预览
        self.update_split_preview()
    
    def end_drag(self, event):
        self.dragging = False
        self.update_preview()
    
    def update_split_preview(self):
        """只更新分割预览，不重新渲染原始图像"""
        if not self.doc or self.current_page >= self.total_pages:
            return
            
        # 获取当前页面
        page = self.doc[self.current_page]
        
        # 清除分割预览画布
        self.left_canvas.delete("all")
        self.right_canvas.delete("all")
        
        # 获取原始图像
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        img_data = pix.tobytes("ppm")
        img = Image.open(io.BytesIO(img_data))
        
        # 创建分割后的图像
        left_img = img.crop((0, 0, int(img.width * self.split_ratio), img.height))
        right_img = img.crop((int(img.width * self.split_ratio), 0, img.width, img.height))
        
        # 调整分割后图像大小以适应画布
        split_canvas_width = self.left_canvas.winfo_width()
        split_canvas_height = self.left_canvas.winfo_height()
        
        canvas_ratio = split_canvas_width / split_canvas_height
        
        left_ratio = left_img.width / left_img.height
        right_ratio = right_img.width / right_img.height
        
        if left_ratio > canvas_ratio:
            left_new_width = split_canvas_width
            left_new_height = int(split_canvas_width / left_ratio)
        else:
            left_new_height = split_canvas_height
            left_new_width = int(split_canvas_height * left_ratio)
        
        if right_ratio > canvas_ratio:
            right_new_width = split_canvas_width
            right_new_height = int(split_canvas_width / right_ratio)
        else:
            right_new_height = split_canvas_height
            right_new_width = int(split_canvas_height * right_ratio)
        
        left_img_resized = left_img.resize((left_new_width, left_new_height), Image.LANCZOS)
        right_img_resized = right_img.resize((right_new_width, right_new_height), Image.LANCZOS)
        
        self.left_img = ImageTk.PhotoImage(left_img_resized)
        self.right_img = ImageTk.PhotoImage(right_img_resized)
        
        # 在分割画布上显示图像
        self.left_canvas.create_image(
            split_canvas_width // 2, split_canvas_height // 2,
            image=self.left_img, anchor=tk.CENTER
        )
        
        self.right_canvas.create_image(
            split_canvas_width // 2, split_canvas_height // 2,
            image=self.right_img, anchor=tk.CENTER
        )
    
    def prev_page(self):
        if self.doc and self.current_page > 0:
            self.current_page -= 1
            self.page_label.config(text=f"{self.current_page + 1}/{self.total_pages}")
            self.update_preview()
    
    def next_page(self):
        if self.doc and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.page_label.config(text=f"{self.current_page + 1}/{self.total_pages}")
            self.update_preview()
    
    def split_and_save(self):
        if not self.doc:
            messagebox.showinfo("提示", "请先选择PDF文件")
            return
        
        output_path = filedialog.asksaveasfilename(
            title="保存分割后的PDF",
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if not output_path:
            return
        
        try:
            # 创建一个新的PDF写入器
            pdf_writer = PdfWriter()
            
            # 清理之前的临时文件
            self.clean_temp_files()
            self.temp_files = []
            
            # 显示进度条
            progress_window = tk.Toplevel(self.root)
            progress_window.title("处理中")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            ttk.Label(progress_window, text="正在分割PDF页面...").pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=self.total_pages)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 处理每一页
            for page_num in range(self.total_pages):
                # 更新进度条
                progress_var.set(page_num)
                progress_window.update()
                
                # 获取当前页面
                page = self.doc[page_num]
                
                # 计算分割位置
                width = page.rect.width
                height = page.rect.height
                split_x = width * self.split_ratio
                
                # 创建左半部分
                left_rect = fitz.Rect(0, 0, split_x, height)
                left_doc = fitz.open()
                left_page = left_doc.new_page(width=split_x, height=height)
                left_page.show_pdf_page(left_page.rect, self.doc, page_num, clip=left_rect)
                
                # 创建右半部分
                right_rect = fitz.Rect(split_x, 0, width, height)
                right_doc = fitz.open()
                right_page = right_doc.new_page(width=width-split_x, height=height)
                right_page.show_pdf_page(right_page.rect, self.doc, page_num, clip=right_rect)
                
                # 保存临时文件
                left_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                right_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                left_temp_name = left_temp.name
                right_temp_name = right_temp.name
                left_temp.close()
                right_temp.close()
                
                left_doc.save(left_temp_name)
                right_doc.save(right_temp_name)
                self.temp_files.extend([left_temp_name, right_temp_name])
                
                # 关闭临时文档
                left_doc.close()
                right_doc.close()
                
                # 添加到输出PDF
                try:
                    left_reader = PdfReader(left_temp_name)
                    right_reader = PdfReader(right_temp_name)
                    
                    pdf_writer.add_page(left_reader.pages[0])
                    pdf_writer.add_page(right_reader.pages[0])
                except Exception as e:
                    messagebox.showerror("错误", f"处理页面 {page_num+1} 时出错: {str(e)}")
                    progress_window.destroy()
                    return
            
            # 保存最终PDF
            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)
            
            # 关闭进度窗口
            progress_window.destroy()
            
            # 清理临时文件
            self.clean_temp_files()
            
            messagebox.showinfo("成功", f"PDF已成功分割并保存到:\n{output_path}")
            self.status_var.set(f"已保存: {os.path.basename(output_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"分割PDF时出错: {str(e)}")
            self.status_var.set("分割PDF失败")
    
    def clean_temp_files(self):
        """安全清理临时文件"""
        for temp_file in self.temp_files:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception as e:
                    print(f"无法删除临时文件 {temp_file}: {str(e)}")
                    # 如果无法立即删除，标记为退出时删除
                    try:
                        import atexit
                        atexit.register(lambda file=temp_file: os.remove(file) if os.path.exists(file) else None)
                    except:
                        pass
    
    def on_closing(self):
        # 清理临时文件
        if self.doc:
            self.doc.close()
            
        self.clean_temp_files()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFA3ToA4Splitter(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
