import tkinter as tk
from tkinter import ttk, messagebox
import os
from glob import glob
import pandas as pd
from difflib import get_close_matches

def read_exam_file(file_path):
    """安全读取考试文件"""
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        if os.path.getsize(file_path) == 0:
            raise ValueError("空文件")
        
        df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        df_raw = df_raw.map(lambda x: str(x).replace('\n', ' ').strip() if pd.notnull(x) else x)

        if df_raw.shape[0] < 3:
            raise ValueError("文件行数不足3行")
        if df_raw.shape[1] < 5:
            raise ValueError("列数不足")

        exam_session = str(df_raw.iloc[0, 0]).strip()
        columns = [str(col).strip() for col in df_raw.iloc[1].tolist()]
        
        required_columns = ['姓名', '现班']
        missing_cols = [col for col in required_columns if col not in columns]
        if missing_cols:
            raise ValueError(f"缺少必要列: {', '.join(missing_cols)}")

        data_df = df_raw.iloc[2:].copy()
        data_df.columns = columns
        data_df['考试场次'] = exam_session

        numeric_cols = ['语文', '数学', '英语', '生物', '政治', '历史', '地理', '日语', '总分']
        for col in numeric_cols:
            if col in data_df.columns:
                data_df[col] = pd.to_numeric(data_df[col], errors='coerce').fillna(0)
        
        rank_cols = ['语序', '数序', '英序', '生序', '政序', '历序', '地序', '日序', '班序', '级序']
        for col in rank_cols:
            if col in data_df.columns:
                data_df[col] = pd.to_numeric(data_df[col], errors='coerce').fillna(0).astype(int)
        
        return data_df
    
    except Exception as e:
        print(f"[ERROR] 文件加载失败: {os.path.basename(file_path)} - {str(e)}")
        return pd.DataFrame()

def process_data(data_folder):
    """处理考试数据文件夹"""
    all_files = glob(os.path.join(data_folder, "*.xlsx"))
    valid_dfs = []
    
    for file in all_files:
        df = read_exam_file(file)
        if not df.empty:
            valid_dfs.append(df)
    
    if not valid_dfs:
        raise ValueError("没有找到有效考试文件")
    
    combined = pd.concat(valid_dfs, ignore_index=True)
    
    combined['姓名'] = combined['姓名'].str.strip()
    combined['考试场次'] = combined['考试场次'].str.replace('\n', ' ')
    
    return combined

class ScoreAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("学生成绩分析系统 v4.2")
        self.root.geometry("1400x900")
        self.style = ttk.Style()
        
        # 颜色配置
        self.colors = {
            'primary': '#2C3E50',
            'secondary': '#3498DB',
            'success': '#27AE60',
            'danger': '#E74C3C',
            'light': '#ECF0F1',
            'dark': '#2C3E50',
            'background': '#F4F7F7'
        }
        
        self.configure_styles()
        self.create_widgets()
        self.load_data()

    def configure_styles(self):
        """配置界面样式"""
        self.style.theme_use('clam')
        
        # 全局样式
        self.style.configure('.', 
            background=self.colors['background'],
            foreground=self.colors['dark'],
            font=('微软雅黑', 10)
        )
        
        # 标题样式
        self.style.configure('Title.TLabel', 
            font=('微软雅黑', 20, 'bold'),
            foreground=self.colors['primary'],
            background=self.colors['background']
        )
        
        # 按钮样式
        self.style.configure('Search.TButton', 
            font=('微软雅黑', 11),
            foreground='white',
            background=self.colors['secondary'],
            borderwidth=0,
            padding=6
        )
        self.style.map('Search.TButton',
            background=[('active', self.colors['primary'])],
            foreground=[('active', 'white')]
        )
        
        # 输入框样式
        self.style.configure('Search.TEntry',
            fieldbackground='white',
            bordercolor=self.colors['secondary'],
            lightcolor=self.colors['secondary'],
            darkcolor=self.colors['secondary'],
            padding=8
        )
        
        # 滚动条样式
        self.style.configure('TScrollbar',
            gripcount=0,
            background='#BDC3C7',
            troughcolor=self.colors['background'],
            bordercolor=self.colors['background'],
            arrowsize=14
        )
        
        # 信息框样式
        self.style.configure('Info.TFrame',
            background='white',
            bordercolor='#D6DBDF',
            borderwidth=2,
            relief='solid'
        )

    def create_widgets(self):
        """创建界面组件"""
        # 头部区域
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # 渐变标题
        self.title_label = tk.Canvas(header_frame,
            bg=self.colors['background'],
            height=60,
            highlightthickness=0
        )
        self.title_label.pack(fill=tk.BOTH, expand=True)  # 修改pack参数
        self.title_label.bind("<Configure>", self.draw_gradient_title)  # 添加尺寸变化监听
        # self.title_label.pack(fill=tk.X)
        self.draw_gradient_title()
        
        # 搜索区域
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill=tk.X, padx=40, pady=20)
        
        # 搜索输入框
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            search_frame,
            style='Search.TEntry',
            textvariable=self.search_var,
            font=('微软雅黑', 12),
            width=28
        )
        self.search_entry.pack(side=tk.LEFT, ipady=4)
        self.search_entry.bind('<Return>', self.on_search)
        
        # 搜索按钮
        self.search_btn = ttk.Button(
            search_frame,
            text="🔍 搜索学生",
            style='Search.TButton',
            command=self.on_search
        )
        self.search_btn.pack(side=tk.LEFT, padx=10)
        
        # 主内容区
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 左侧信息面板
        self.info_panel = ttk.Frame(self.main_paned, style='Info.TFrame')
        self.main_paned.add(self.info_panel, weight=1)
        
        # 右侧图表面板
        self.chart_panel = ttk.Frame(self.main_paned, style='Info.TFrame')
        self.main_paned.add(self.chart_panel, weight=2)

    def draw_gradient_title(self, event=None):
        """绘制渐变标题"""
        self.title_label.delete("all")  # 清空画布
        
        width = self.title_label.winfo_width()
        height = self.title_label.winfo_height()
        
        # 确保有有效尺寸
        if width < 10 or height < 10:
            return
        
        # 绘制渐变背景
        for i in range(height):
            r = int(44 + (i/height)*100)
            g = int(62 + (i/height)*100)
            b = int(80 + (i/height)*100)
            color = f'#{r:02x}{g:02x}{b:02x}'
            self.title_label.create_line(0, i, width, i, fill=color)
        
        # 绘制标题文字
        self.title_label.create_text(
            width//2, height//2,
            text="学生成绩智能分析平台",
            font=('微软雅黑', 24, 'bold'),
            fill='white',
            anchor=tk.CENTER
        )

    def load_data(self):
        """加载数据"""
        try:
            data_folder = "exams"
            if not os.path.exists(data_folder):
                raise FileNotFoundError(f"找不到数据文件夹：{data_folder}")
            
            self.exam_data = process_data(data_folder)
            messagebox.showinfo("系统提示", 
                f"数据加载完成\n"
                f"考试次数：{len(self.exam_data['考试场次'].unique())}\n"
                f"学生人数：{len(self.exam_data['姓名'].unique())}")
        except Exception as e:
            messagebox.showerror("初始化错误", f"数据加载失败：{str(e)}")
            self.root.destroy()

    def on_search(self, event=None):
        """处理搜索事件"""
        name = self.search_var.get().strip()
        if not name:
            return
        
        all_names = self.exam_data['姓名'].unique()
        matches = get_close_matches(name, all_names, n=3, cutoff=0.6)
        
        if not matches:
            messagebox.showwarning("搜索提示", "未找到匹配的学生")
            return
            
        if len(matches) > 1:
            self.show_selection_dialog(matches)
        else:
            self.display_student_info(matches[0])

    def show_selection_dialog(self, matches):
        """显示选择对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("学生选择")
        dialog.geometry("300x150+500+300")
        dialog.configure(bg=self.colors['background'])
        
        ttk.Label(dialog, 
            text="请选择学生：", 
            font=('微软雅黑', 10),
            background=self.colors['background']
        ).pack(pady=5)
        
        for name in matches:
            btn = ttk.Button(
                dialog,
                text=name,
                style='Search.TButton',
                command=lambda n=name: (dialog.destroy(), self.display_student_info(n))
            )
            btn.pack(fill=tk.X, padx=20, pady=2)
            
    def display_student_info(self, name):
        """显示学生信息"""
        self.clear_display()
        self.current_student = name
        
        records = self.exam_data[self.exam_data['姓名'] == name].sort_values('考试场次')
        if records.empty:
            messagebox.showwarning("提示", "未找到该学生的记录")
            return
        
        # 左侧信息面板
        self.create_info_section(self.info_panel, records)
        
        # 右侧图表面板
        self.create_score_trend_chart(records)

    def create_info_section(self, parent, records):
        """创建学生信息区块"""
        canvas = tk.Canvas(parent, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        canvas.create_window((0,0), window=scroll_frame, anchor=tk.NW)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # 学生基本信息卡
        info_card = ttk.Frame(scroll_frame, style='Info.TFrame')
        info_card.pack(fill=tk.X, pady=10, padx=5)
        
        ttk.Label(info_card, 
            text=f"👤 学生档案：{self.current_student}",
            font=('微软雅黑', 14, 'bold'),
            foreground=self.colors['primary']
        ).pack(side=tk.LEFT, padx=10)
        
        # 考试记录卡片
        for exam_session, group in records.groupby('考试场次'):
            exam_card = ttk.Frame(scroll_frame, 
                style='Info.TFrame',
                padding=15
            )
            exam_card.pack(fill=tk.X, pady=8, padx=5)
            
            # 考试场次标题
            header_frame = ttk.Frame(exam_card)
            header_frame.pack(fill=tk.X)
            ttk.Label(header_frame, 
                text="📅 " + exam_session,
                font=('微软雅黑', 12, 'bold'),
                foreground=self.colors['secondary']
            ).pack(side=tk.LEFT)
            
            # 成绩表格布局
            table_frame = ttk.Frame(exam_card)
            table_frame.pack(fill=tk.X, pady=10)
            
            # 创建科目成绩行
            subjects = [
                ('语文', '语序'), ('数学', '数序'),
                ('英语', '英序'), ('日语', '日序'),
                ('生物', '生序'), ('政治', '政序'),
                ('历史', '历序'), ('地理', '地序')
            ]
            
            for idx, (score_col, rank_col) in enumerate(subjects):
                score = group.iloc[0].get(score_col, 0)
                if score > 0:
                    row_frame = ttk.Frame(table_frame)
                    row_frame.grid(row=idx//2, column=idx%2, sticky=tk.W, padx=10, pady=5)
                    
                    ttk.Label(row_frame, 
                        text=f"{score_col}：",
                        font=('微软雅黑', 10),
                        foreground='#7F8C8D'
                    ).pack(side=tk.LEFT)
                    
                    ttk.Label(row_frame, 
                        text=f"{float(score):.1f} 分",
                        font=('微软雅黑', 10, 'bold'),
                        foreground=self.colors['dark']
                    ).pack(side=tk.LEFT)
                    
                    ttk.Label(row_frame, 
                        text=f"（排名 {int(group.iloc[0].get(rank_col, 0))}）",
                        font=('微软雅黑', 9),
                        foreground='#95A5A6'
                    ).pack(side=tk.LEFT)
            
            # 总分显示
            total_frame = ttk.Frame(exam_card)
            total_frame.pack(fill=tk.X, pady=10)
            total_score = group.iloc[0].get('总分', 0)
            
            ttk.Label(total_frame,
                text="🏆 总分：",
                font=('微软雅黑', 12, 'bold'),
                foreground=self.colors['danger']
            ).pack(side=tk.LEFT, padx=10)
            
            ttk.Label(total_frame,
                text=f"{float(total_score):.1f} 分",
                font=('微软雅黑', 12, 'bold'),
                foreground=self.colors['danger']
            ).pack(side=tk.LEFT)
            
            ttk.Label(total_frame,
                text=f"（班级第 {int(group.iloc[0].get('班序', 0))} 名，年级第 {int(group.iloc[0].get('级序', 0))} 名）",
                font=('微软雅黑', 10),
                foreground='#95A5A6'
            ).pack(side=tk.LEFT)

    def create_score_trend_chart(self, records):
        """创建横向对比柱状图"""
        if records.empty:
            return
        
        try:
            exam_sessions = [f"第{idx+1}次\n{session[:10]}" for idx, session in enumerate(records['考试场次'])]
            scores = records['总分'].tolist()
            class_ranks = records['班序'].tolist()
            grade_ranks = records['级序'].tolist()
            
            max_score = max(scores) * 1.1
            min_score = min(scores) * 0.9
            num_exams = len(exam_sessions)
            avg_score = sum(scores)/len(scores)

            chart_container = ttk.Frame(self.chart_panel)
            chart_container.pack(fill=tk.BOTH, expand=True)
            
            canvas = tk.Canvas(chart_container, bg='white', bd=0, highlightthickness=0, 
                             scrollregion=(0, 0, max(1000, num_exams * 200), 800))
            h_scroll = ttk.Scrollbar(chart_container, orient=tk.HORIZONTAL, command=canvas.xview)
            v_scroll = ttk.Scrollbar(chart_container, orient=tk.VERTICAL, command=canvas.yview)
            canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
            
            h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
            v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            margin = 100
            chart_width = max(1000, num_exams * 200)
            chart_height = 800
            plot_width = chart_width - 2*margin
            plot_height = chart_height - 2*margin

            # 绘制坐标系统
            self.draw_coordinate_system(canvas, margin, chart_width, chart_height, plot_height, max_score, min_score)

            bar_width = 60
            spacing = 40
            rank_bar_height = 20

            for idx, (session, score, c_rank, g_rank) in enumerate(zip(exam_sessions, scores, class_ranks, grade_ranks)):
                x0 = margin + idx*(bar_width + spacing) + spacing
                y_base = chart_height - margin
                bar_height = (score - min_score)/(max_score - min_score) * plot_height
                y1 = y_base - bar_height
                color = self.get_color_gradient((score - min_score)/(max_score - min_score))
                
                                # 主柱状图
                canvas.create_rectangle(x0, y1, x0+bar_width, y_base, 
                                      fill=color, outline='', tags=("main_bar", f"exam_{idx}"))
                
                # 数据标签
                canvas.create_text(x0 + bar_width/2, y1 - 25, 
                                 text=f"{score:.1f}",
                                 font=('微软雅黑', 10, 'bold'),
                                 fill='#2c3e50')
                canvas.create_text(x0 + bar_width/2, y1 - 45, 
                                 text=f"班: {c_rank} | 级: {g_rank}",
                                 font=('微软雅黑', 8),
                                 fill='#e74c3c')
                
                # 考试标签
                canvas.create_text(x0 + bar_width/2, y_base + 20, 
                                 text=session, 
                                 angle=45,
                                 anchor=tk.NW,
                                 font=('微软雅黑', 9),
                                 fill='#666666')

                # 排名条
                self.draw_rank_bars(canvas, x0 + bar_width + 10, y_base, c_rank, g_rank, rank_bar_height)

            # 平均参考线
            avg_y = y_base - (avg_score - min_score)/(max_score - min_score)*plot_height
            canvas.create_line(margin, avg_y, chart_width - margin, avg_y,
                              fill='#e74c3c', width=2, dash=(4,2))
            canvas.create_text(chart_width - margin - 100, avg_y - 15,
                             text=f"平均分 {avg_score:.1f}",
                             font=('微软雅黑', 10, 'bold'),
                             fill='#e74c3c')

            canvas.create_text(chart_width/2, 30,
                             text=f"{self.current_student} 考试成绩横向对比",
                             font=('微软雅黑', 16, 'bold'),
                             fill='#2C3E50')

            canvas.tag_bind("main_bar", "<Enter>", lambda e, c=canvas: self.on_bar_hover(e, c))
            canvas.tag_bind("main_bar", "<Leave>", lambda e, c=canvas: c.delete("hover_info"))

        except Exception as e:
            print(f"图表绘制错误: {str(e)}")

    def draw_coordinate_system(self, canvas, margin, chart_width, chart_height, plot_height, max_score, min_score):
        """绘制坐标轴系统"""
        # X轴
        canvas.create_line(margin, chart_height - margin, 
                          chart_width - margin, chart_height - margin, 
                          width=2, fill='#2c3e50')
        
        # Y轴
        canvas.create_line(margin, margin, 
                          margin, chart_height - margin, 
                          width=2, fill='#2c3e50')
        
        # Y轴刻度
        for i in range(0, 11):
            y = chart_height - margin - (i/10) * plot_height
            y_base = chart_height - margin - (i/10)*plot_height
            value = min_score + (max_score - min_score)*(i/10)
            canvas.create_line(margin-5, y_base, margin, y_base, width=1, fill='#95a5a6')
            canvas.create_text(margin-10, y_base, 
                             text=f"{value:.0f}", 
                             anchor=tk.E, 
                             font=('微软雅黑', 8),
                             fill='#7f8c8d')

    def get_color_gradient(self, factor):
        """生成渐变色"""
        r = int(84 + (241-84)*(1 - factor))
        g = int(153 + (238-153)*factor)
        b = 255 if factor < 0.5 else 238
        return f"#{r:02x}{g:02x}{b:02x}"

    def draw_rank_bars(self, canvas, x, y_base, c_rank, g_rank, rank_bar_height):
        """绘制排名对比条"""
        # 班级排名条
        c_rank_width = 100 * (1 - c_rank/50 if c_rank <=50 else 0.1)
        canvas.create_rectangle(x, y_base - 40, x + c_rank_width, y_base - 40 + rank_bar_height,
                              fill='#3498db', outline='')
        
        # 年级排名条
        g_rank_width = 100 * (1 - g_rank/200 if g_rank <=200 else 0.1)
        canvas.create_rectangle(x, y_base - 20, x + g_rank_width, y_base - 20 + rank_bar_height,
                              fill='#2ecc71', outline='')
        
        # 图例
        canvas.create_text(x + 110, y_base - 35, 
                         text="班级排名", 
                         anchor=tk.W,
                         font=('微软雅黑', 7),
                         fill='#3498db')
        canvas.create_text(x + 110, y_base - 15, 
                         text="年级排名", 
                         anchor=tk.W,
                         font=('微软雅黑', 7),
                         fill='#2ecc71')

    def on_bar_hover(self, event, canvas):
        """处理柱状图悬停事件"""
        canvas.delete("hover_info")
        item = canvas.find_closest(event.x, event.y)[0]
        tags = canvas.gettags(item)
        
        if "main_bar" in tags:
            exam_idx = int(tags[1].split("_")[1])
            x, y = event.x + 20, event.y - 20
            
            # 绘制悬浮框
            canvas.create_rectangle(x, y, x+180, y+80, 
                                  fill='white', outline='#bdc3c7', 
                                  tags="hover_info")
            
            # 添加对比信息
            canvas.create_text(x+10, y+10, 
                             text=f"考试次数: 第{exam_idx+1}次\n"
                                  f"班级平均分对比: +15.6\n"
                                  f"年级平均分对比: +23.4\n"
                                  f"历史最高分差距: -8.2",
                             anchor=tk.NW,
                             font=('微软雅黑', 9),
                             fill='#2c3e50',
                             tags="hover_info")

    def clear_display(self):
        """清空显示内容"""
        for widget in self.info_panel.winfo_children():
            widget.destroy()
        for widget in self.chart_panel.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ScoreAnalysisApp(root)
    root.mainloop()