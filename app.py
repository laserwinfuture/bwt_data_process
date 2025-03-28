import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np
from matplotlib.font_manager import FontProperties

# 设置页面配置
st.set_page_config(
    page_title="M2数据分析工具",
    page_icon="📊",
    layout="wide"
)

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

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
    ax1.set_title('M2聚焦光斑分析', fontsize=14)
    ax1.set_xlabel('Z Location (mm)', fontsize=12)
    ax1.set_ylabel('Beam Width (μm)', fontsize=12)
    ax1.legend(loc='upper left', fontsize=10)
    ax1.grid(True)
    
    # 创建第二个Y轴用于比值
    ax2 = ax1.twinx()
    ax2.plot(df.iloc[:, 2], ratio, 'g--', label='光斑圆度')
    
    # 添加合格标准线
    standard_line = ax2.axhline(y=0.9, color='orange', linestyle='-', linewidth=2, alpha=0.7, label='合格标准(0.9)')
    
    # 添加合格区域填充
    x_range = df.iloc[:, 2]
    ax2.fill_between(x_range, 0.9, 1.0, color='lightgreen', alpha=0.3, label='合格区域')
    ax2.fill_between(x_range, 0.85, 0.9, color='pink', alpha=0.3, label='不合格区域')
    
    ax2.set_ylabel('聚焦光斑圆度', color='g', fontsize=12)
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

def main():
    # 侧边栏
    with st.sidebar:
        st.title("功能选择")
        selected_function = st.radio(
            "请选择功能：",
            ["M2数据二次处理"]
        )
    
    # 主页面
    if selected_function == "M2数据二次处理":
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

if __name__ == "__main__":
    main() 