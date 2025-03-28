import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np
import platform
import matplotlib.font_manager as fm
import os
import time
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from decimal import Decimal, ROUND_HALF_UP

__version__ = '0.1'

# 设置中文字体
#plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS', 'sans-serif']  # 用来正常显示中文
#plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号


# 跨平台字体设置函数
def setup_chinese_font():
    # 不同平台的常用中文字体
    if platform.system() == 'Windows':
        font_list = ['SimHei', 'Microsoft YaHei', 'SimSun']
    elif platform.system() == 'Darwin':  # macOS
        font_list = ['PingFang SC', 'Heiti SC', 'STHeiti']
    else:  # Linux
        font_list = ['WenQuanYi Micro Hei', 'Noto Sans CJK SC', 'DejaVu Sans']
    
    # 获取系统可用字体
    available_fonts = [f.name for f in fm.fontManager.ttflist]
    
    # 查找第一个可用的中文字体
    for font in font_list:
        if font in available_fonts:
            plt.rcParams['font.sans-serif'] = [font, 'DejaVu Sans', 'Arial Unicode MS']
            plt.rcParams['axes.unicode_minus'] = False
            return font
    
    # 如果没有找到中文字体，使用默认设置
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    return 'DejaVu Sans'

# 在模块加载时设置字体
setup_chinese_font()



# 设置页面配置
st.set_page_config(
    page_title="M2数据分析工具",
    page_icon="��",
    layout="wide"
)

def process_m2_data(file_content):
    """处理M2数据并生成图表"""
    # 找到Frame (Quantitative)部分
    frame_part = file_content.split('Frame (Quantitative)')[1].strip()
    
    # 使用pandas读取这部分数据
    df = pd.read_csv(io.StringIO(frame_part))
    
    # 计算比值数据
    ratio = []
    for i in range(len(df)):
        x_val = df.iloc[i, 0]  # Beam Width X
        y_val = df.iloc[i, 1]  # Beam Width Y
        if y_val / x_val > 1:
            ratio.append(x_val / y_val)
        else:
            ratio.append(y_val / x_val)
    
    # 创建图形
    fig, ax1 = plt.subplots(figsize=(10, 6))
    
    # 绘制曲线
    ax1.plot(df.iloc[:, 2], df.iloc[:, 0], 'r-', label='Beam Width X (μm)')
    ax1.plot(df.iloc[:, 2], df.iloc[:, 1], 'b-', label='Beam Width Y (μm)')
    
    # 设置第一个Y轴的标签
    ax1.set_title('M2 foucs analysis', fontsize=14)
    ax1.set_xlabel('Z Location (mm)', fontsize=12)
    ax1.set_ylabel('Beam Width (μm)', fontsize=12)
    ax1.legend(loc='upper left', fontsize=10)
    ax1.grid(True)
    
    # 创建第二个Y轴用于比值
    ax2 = ax1.twinx()
    ax2.plot(df.iloc[:, 2], ratio, 'g--', label='beam roundness')
    
    # 添加合格标准线
    standard_line = ax2.axhline(y=0.9, color='orange', linestyle='-', linewidth=2, alpha=0.7, label='Qualification criteria(0.9)')
    
    # 添加合格区域填充
    x_range = df.iloc[:, 2]
    ax2.fill_between(x_range, 0.9, 1.0, color='lightgreen', alpha=0.3, label='qualified area')
    ax2.fill_between(x_range, 0.85, 0.9, color='pink', alpha=0.3, label='unqualified area')
    
    ax2.set_ylabel('beam roundness(after focus)', color='g', fontsize=12)
    ax2.tick_params(axis='y', colors='g')
    ax2.set_ylim(0.85, 1.0)
    ax2.legend(loc='upper right', fontsize=10)
    
    plt.tight_layout()
    return fig

def save_fig_to_bytes(fig):
    """将matplotlib图形保存为字节流"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    buf.seek(0)
    return buf

def process_summary_data(template_file, data_files):
    """处理纠正预防措施汇总数据"""
    try:
        # 读取模板文件
        wb = load_workbook(template_file)
        ws = wb.active
        
        # 获取第四行的数据作为映射表
        first_row_data = {}
        for col in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col)
            cell_value = ws.cell(row=4, column=col).value
            if cell_value is not None:
                first_row_data[col_letter] = cell_value
        
        # 从第5行开始处理数据文件
        row = 5
        processed_files_count = 0
        
        for data_file in data_files:
            load_wb = load_workbook(data_file, data_only=True)
            load_sheet = load_wb.active
            
            # 写入序号和文件名
            ws[f'A{row}'] = row-4
            ws[f'B{row}'] = os.path.basename(data_file.name)
            
            # 根据映射表处理数据
            for target_cell, source_cell in first_row_data.items():
                if source_cell is not None:
                    source_value = load_sheet[source_cell].value
                    if isinstance(source_value, datetime):
                        source_value = source_value.strftime('%Y/%m/%d')
                    ws[f'{target_cell}{row}'] = source_value
            
            processed_files_count += 1
            row += 1
            load_wb.close()
        
        # 保存处理后的文件
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer, processed_files_count
        
    except Exception as e:
        st.error(f"处理文件时出错：{str(e)}")
        return None, 0

def round_to_decimal(value: str, decimal_places: int) -> float:
    """
    将字符串数值四舍五入到指定小数位数
    Args:
        value (str): 需要四舍五入的字符串数值
        decimal_places (int): 四舍五入后保留的小数位数
    Returns:
        float: 四舍五入后的浮点数
    """
    format_str = '0.' + '0' * decimal_places
    return float(Decimal(value).quantize(Decimal(format_str), rounding=ROUND_HALF_UP))

def process_product_data(template_file, data_files):
    """处理产成品数据汇总"""
    try:
        # 读取模板文件
        wb = load_workbook(template_file)
        ws = wb.active
        
        # 获取第一行的数据作为映射表
        first_row_data = {}
        for col in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col)
            cell_value = ws.cell(row=1, column=col).value
            if cell_value is not None:
                first_row_data[col_letter] = cell_value
        
        # 从最后一行开始处理数据文件
        row = ws.max_row + 1
        processed_files_count = 0
        
        for data_file in data_files:
            load_wb = load_workbook(data_file, data_only=True)
            load_sheet = load_wb.active
            
            # 写入序号和文件名
            ws[f'A{row}'] = row-1
            ws[f'B{row}'] = os.path.basename(data_file.name)
            
            # 根据映射表处理数据
            for target_cell, source_cell in first_row_data.items():
                if source_cell is not None:
                    source_value = load_sheet[source_cell].value
                    cell = ws[f'{target_cell}{row}']
                    cell.value = source_value
                    
                    # 如果是日期类型，设置日期格式
                    if isinstance(source_value, datetime):
                        cell.number_format = 'yyyy/mm/dd'
                    
                    # 特殊处理AR列（功率范围）
                    if target_cell == 'AR':
                        var1, var2 = source_value.split('-')
                        var1 = int(var1[:-3])
                        var2 = int(var2[:-5])
                        ws[f'AR{row}'] = var1
                        ws[f'AS{row}'] = var2
                    
                    # 特殊处理AT列（功率计算）
                    elif target_cell == 'AT':
                        power = ws[f'AT{row}'].value
                        ws[f'AU{row}'] = round_to_decimal(power / var1 * 1000, 0)
            
            processed_files_count += 1
            row += 1
            load_wb.close()
        
        # 添加边框
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        # 只为新添加的行添加边框
        for row in range(ws.max_row - processed_files_count + 1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).border = thin_border
        
        # 保存处理后的文件
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer, processed_files_count
        
    except Exception as e:
        st.error(f"处理文件时出错：{str(e)}")
        return None, 0

def main():
    # 侧边栏
    with st.sidebar:
        st.title("功能选择")
        
        # 设置按钮样式
        st.markdown("""
            <style>
            .stButton > button {
                font-size: 3rem;
                padding: 1.2rem;
                margin: 0.8rem 0;
                border-radius: 10px;
                background-color: #7832E1;
                color: white;
                border: none;
                transition: all 0.3s ease;
            }
            .stButton > button:hover {
                background-color: #6B2BC7;
                transform: translateY(-2px);
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            }
            </style>
        """, unsafe_allow_html=True)
        
        # 使用按钮进行功能切换
        if st.button("M2数据二次处理", use_container_width=True):
            st.session_state.selected_function = "M2数据二次处理"
        if st.button("纠正预防措施汇总", use_container_width=True):
            st.session_state.selected_function = "纠正预防措施汇总"
        if st.button("产成品数据汇总", use_container_width=True):
            st.session_state.selected_function = "产成品数据汇总"
        
        # 初始化选择的功能
        if 'selected_function' not in st.session_state:
            st.session_state.selected_function = "M2数据二次处理"
    
    # 主页面
    if st.session_state.selected_function == "M2数据二次处理":
        st.title("M2数据二次处理")
        st.write("请上传CSV文件进行处理")
        
        # 文件上传
        uploaded_file = st.file_uploader("选择或拖拽CSV文件", type=['csv'])
        
        if uploaded_file is not None:
            try:
                # 读取文件内容
                file_content = uploaded_file.getvalue().decode('utf-8')
                
                # 处理数据并显示图表
                fig = process_m2_data(file_content)
                st.pyplot(fig)
                
                # 添加下载按钮
                img_bytes = save_fig_to_bytes(fig)
                st.download_button(
                    label="下载图表",
                    data=img_bytes,
                    file_name="beam_analysis.png",
                    mime="image/png"
                )
                
            except Exception as e:
                st.error(f"处理文件时出错：{str(e)}")
                st.write("请确保上传的文件格式正确，包含'Frame (Quantitative)'部分")
    
    elif st.session_state.selected_function == "纠正预防措施汇总":
        st.title("纠正预防措施汇总")
        
        # 模板文件上传
        st.subheader("1. 上传模板文件")
        template_file = st.file_uploader("请上传模板文件（Excel格式）", type=['xlsx'], key="template")
        
        # 数据文件上传
        st.subheader("2. 上传需要汇总的文件")
        data_files = st.file_uploader("请上传需要汇总的文件（Excel格式）", type=['xlsx'], accept_multiple_files=True, key="data")
        
        if template_file and data_files:
            if st.button("开始处理"):
                with st.spinner("正在处理文件..."):
                    start_time = time.time()
                    
                    # 处理数据
                    output_buffer, processed_count = process_summary_data(template_file, data_files)
                    
                    if output_buffer:
                        # 显示处理结果
                        st.subheader("3. 处理结果")
                        st.success(f"成功处理 {processed_count} 个文件")
                        
                        # 显示处理时间
                        elapsed_time = time.time() - start_time
                        st.info(f"处理用时: {elapsed_time:.2f} 秒")
                        
                        # 提供下载按钮
                        st.download_button(
                            label="下载汇总文件",
                            data=output_buffer,
                            file_name="纠正预防措施汇总表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("处理失败，请检查文件格式是否正确")
    
    elif st.session_state.selected_function == "产成品数据汇总":
        st.title("产成品数据汇总")
        
        # 模板文件上传
        st.subheader("1. 上传模板文件")
        template_file = st.file_uploader("请上传模板文件（Excel格式）", type=['xlsx'], key="product_template")
        
        # 数据文件上传
        st.subheader("2. 上传需要汇总的文件")
        data_files = st.file_uploader("请上传需要汇总的文件（Excel格式）", type=['xlsx'], accept_multiple_files=True, key="product_data")
        
        if template_file and data_files:
            if st.button("开始处理"):
                with st.spinner("正在处理文件..."):
                    start_time = time.time()
                    
                    # 处理数据
                    output_buffer, processed_count = process_product_data(template_file, data_files)
                    
                    if output_buffer:
                        # 显示处理结果
                        st.subheader("3. 处理结果")
                        st.success(f"成功处理 {processed_count} 个文件")
                        
                        # 显示处理时间
                        elapsed_time = time.time() - start_time
                        st.info(f"处理用时: {elapsed_time:.2f} 秒")
                        
                        # 提供下载按钮
                        st.download_button(
                            label="下载汇总文件",
                            data=output_buffer,
                            file_name="产成品数据汇总表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("处理失败，请检查文件格式是否正确")

if __name__ == "__main__":
    main()