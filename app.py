import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openai import OpenAI
import json
from datetime import datetime
import textwrap
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from PIL import Image
import io
import base64
import numpy as np
from collections import Counter
import pdfkit
from jinja2 import Template
import os
import time
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO
from PIL import Image as PILImage
import plotly.io as pio

# Configure page settings
st.set_page_config(
    page_title="Yakjin - Brand Analysis",
    page_icon="👕",
    layout="wide"
)

# CSS 스타일 수정
st.markdown("""
    <style>
        .chart-title {
            font-size: 1.53rem;
            font-weight: bold;
            color: #333;
            margin-bottom: 1rem;
            text-align: center;
            width: 100%;
            display: block;
        }
        
        .chart-container {
            margin-bottom: 2.5rem;
        }
        
        .design-overview {
            font-size: 1.3rem;
            font-weight: bold;
            margin-bottom: 1.5rem;
        }
        
        /* 제목 래 실선 제거 */
        h3 {
            border-bottom: none !important;
        }
        
        /* 실선 스타일 */
        hr {
            border: none;
            height: 1px;
            background-color: #e6e6e6;
            margin-top: 2rem;
        }

        /* 새로운 레이아웃 스타일 추가 */
        .block-container {
            padding-left: 400px !important;
            padding-right: 400px !important;
            max-width: 100% !important;
        }

        /* 모바일 대응을 위한 미디어 쿼리 */
        @media (max-width: 1200px) {
            .block-container {
                padding-left: 50px !important;
                padding-right: 50px !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

# Initialize OpenAI client
client = OpenAI(api_key="sk-q2r6kBRqi4fiE7N9nFtEVV8xk79ymwPyG7ee8I-QFET3BlbkFJ2xvRuSukBKf37ZOkKAQuFxR5fY5IasxIHoPtajS5sA")

# Color mapping definition
color_mapping = {
    # 기본 색상 (White를 연한 회색으로 변경)
    'Black': '#2C2C2C',      # 순수 검정 대신 부드러운 차콜
    'White': '#F2F2F2',      # 순수 흰색 대신 더 진한 회색빛으로 변경
    'Gray': '#B4B4B4',       # 중간 회색
    
    # 파스텔 블루 계열
    'Navy': '#7B89B3',       # 파스텔 네이비
    'Blue': '#BAE1FF',       # 파스텔 블루
    'Sky Blue': '#D4F1F9',   # 파스텔 스카이블루
    'Mint': '#B5EAD7',       # 파스텔 민트
    
    # 파스텔 레드/핑크 계열
    'Red': '#FFB3B3',        # 파스텔 레드
    'Pink': '#FEC8D8',       # 파스텔 핑크
    'Coral': '#FFB5A7',      # 파스텔 코랄
    'Rose': '#F8C4D4',       # 파스텔 로즈
    
    # 파스텔 퍼플 계열
    'Purple': '#E0BBE4',     # 파스텔 퍼플
    'Lavender': '#D4BBEF',   # 파스텔 라벤더
    'Mauve': '#E2BCD6',      # 파스텔 모브
    
    # 파스텔 그린 계열
    'Green': '#BAFFC9',      # 파스텔 그린
    'Sage': '#CCE2CB',       # 파스텔 세이지
    'Olive': '#D1E0BF',      # 파스텔 올리브
    
    # 파스텔 옐로우/브라운 계열
    'Yellow': '#FFE4BA',     # 파스텔 옐로우
    'Beige': '#FFDFD3',      # 파스텔 베이지
    'Brown': '#E6C9A8',      # 파스텔 브라운
    'Camel': '#E6CCB2',      # 파스텔 카멜
    
    # 기타 파스텔 색상
    'Orange': '#FFD4B8',     # 파스텔 오렌지
    'Peach': '#FFDAC1',      # 파스텔 피치
    'Khaki': '#E6D5B8',      # 파스텔 카키
    
    # 기타
    'Multi': '#E5E5E5',      # 멀티컬러 표현
    'Other': '#DDDDDD'       # 기타 색상
}

def detect_file_structure(df):
    """파일 구조를 감지하고 적절한 시작 행을 반환"""
    for idx in range(min(5, len(df))):
        if df.iloc[idx].astype(str).str.contains('Title|Product|Item', case=False).any():
            return idx
    return 0

def get_required_columns():
    """분석에 필요한 최소 필수 컬럼 정의"""
    return {
        'Title': str,
        'Category': str,
        'Original Price (USD)': float,
        'Current Price (USD)': float,
        'Materials': str,
        'Discount': float
    }

def clean_dataframe(df):
    """데이터프레임 정제 및 준비"""
    try:
        # 입력 데이터프레임 복사본 생성
        df = df.copy()
        
        # Discount 값이 소수점(0.2)인지 퍼센트(20)인지 확인하 통일
        if df['Discount'].mean() <= 1:  # 소수점 형태(0.2)라면
            df['Discount'] = df['Discount'] * 100  # 퍼센트로 변환
        
        # 나머지 숫자형 컬럼 정제
        df['Original Price (USD)'] = pd.to_numeric(df['Original Price (USD)'], errors='coerce')
        df['Current Price (USD)'] = pd.to_numeric(df['Current Price (USD)'], errors='coerce')
        
        # 문자형 컬럼 정제 - 빈 문자열 대신 의미 있는 값으로 대체
        df['Title'] = df['Title'].fillna('Untitled')
        df['Category'] = df['Category'].fillna('Uncategorized')
        df['Materials'] = df['Materials'].fillna('Not Specified')
        
        # Discount가 NaN인 경우 0으로 처리
        df['Discount'] = df['Discount'].fillna(0)
        
        # 모든 값이 null인 행 제거
        df = df.dropna(how='all')
        
        # 컬럼명 변환
        column_mapping = {
            'Original Price (USD)': 'Original_Price',
            'Current Price (USD)': 'Current_Price'
        }
        df = df.rename(columns=column_mapping)
        
        return df

    except Exception as e:
        st.error(f"데이터 정제 중 오류 발생: {str(e)}")
        st.write("현재 컬럼:", df.columns.tolist())
        return None

def get_ai_insights(data_summary):
    """Get AI-Powered insights from the data"""
    try:
        # 프로그레스 바 추가
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        prompt = f"""
        Analyze the following data and provide insights in this exact format:

        # 📊 AI-Powered Insights

        👕 Product Assortment
        Analysis: [Your analysis of product distribution]
        Suggestion: [Your suggestion for product assortment]

        🧵 Material Composition
        Analysis: [Your analysis of material composition]
        Suggestion: [Your suggestion for materials]

        💰 Price & Discount
        Analysis: [Your analysis of pricing and discounts]
        Suggestion: [Your suggestion for pricing strategy]

        Use actual data values from:
        Category Distribution: {json.dumps(data_summary['product_distribution'])}
        Price Metrics: {json.dumps(data_summary['price_range'])}
        Material Data: {json.dumps(data_summary.get('material_stats', {}))}
        Discount Information: {json.dumps(data_summary['discount_stats'])}
        """
        
        # CSS 스타일 정의
        st.markdown("""
            <style>
                .main-insights-title {
                    font-size: 2.07rem;
                    font-weight: bold;
                    margin-bottom: 1.4rem;
                    color: #333;
                    border-top: 1px solid #e6e6e6;
                    padding-top: 2rem;
                }
                .section-insights-title {
                    font-size: 1.6rem;
                    font-weight: bold;
                    margin-top: 1.5rem;
                    margin-bottom: 1rem;
                    color: #333;
                }
                .analysis-text {
                    font-size: 1rem;
                    margin-bottom: 1rem;
                    color: #333;
                }
                .suggestion-text {
                    font-size: 1rem;
                    margin-bottom: 1.5rem;
                    color: #333;
                }
            </style>
        """, unsafe_allow_html=True)
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": "You are a senior fashion retail analyst. Provide specific, data-driven insights and avoid generic advice. Keep responses concise and focused."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.5,
            max_tokens=300
        )

        # 프로그레스 바 업데이트
        for i in range(100):
            progress_bar.progress(i + 1)
            status_text.text(f"Generating insights... {i + 1}%")
            time.sleep(0.01)

        insights = response.choices[0].message.content
        
        # 프로그레스 바와 상태 텍스트 제거
        progress_bar.empty()
        status_text.empty()

        # 메인 타이틀 표시
        st.markdown('<p class="main-insights-title">📊 AI-Powered Insights</p>', unsafe_allow_html=True)
        
        # GPT 응답 파싱 및 포맷팅 - 한 번만 실행
        sections = insights.split('\n\n')
        for section in sections[1:]:  # 첫 번째 섹션(타이틀) 제외
            lines = section.strip().split('\n')
            if len(lines) >= 1:
                # 섹션 타이틀
                st.markdown(f'<p class="section-insights-title">{lines[0]}</p>', unsafe_allow_html=True)
                
                # Analysis와 Suggestion 파싱
                for line in lines[1:]:
                    if line.startswith('Analysis:'):
                        # Analysis: 부분만 볼드체로 처
                        text = line.replace('Analysis:', '<strong>Analysis:</strong>')
                        st.markdown(f'<p class="analysis-text">{text}</p>', unsafe_allow_html=True)
                    elif line.startswith('Suggestion:'):
                        # Suggestion: 부분만 볼드체로 처리
                        text = line.replace('Suggestion:', '<strong>Suggestion:</strong>')
                        st.markdown(f'<p class="suggestion-text">{text}</p>', unsafe_allow_html=True)
        
        # insights를 반환하지 않음
        return None

    except Exception as e:
        st.error(f"Error generating insights: {str(e)}")
        return None

# 데이터 요약 준비 함수
def prepare_data_summary(df):
    """분석을 위한 데이터 요약 준비"""
    return {
        "product_distribution": {str(k): int(v) for k, v in df['Category'].value_counts().to_dict().items()},
        "price_range": {
            "min": float(df['Current_Price'].min()),
            "max": float(df['Current_Price'].max()),
            "avg": float(df['Current_Price'].mean()),
            "median": float(df['Current_Price'].median())
        },
        "discount_stats": {
            "avg_discount": float(df['Discount'].mean()),
            "max_discount": float(df['Discount'].max()),
            "discount_distribution": {str(k): int(v) for k, v in df['Discount'].value_counts().to_dict().items()}
        },
        "material_stats": {str(k): int(v) for k, v in df['Materials'].value_counts().to_dict().items()},
    }

def analyze_data(df, uploaded_file):
    """Analyze the uploaded data and create visualizations"""
    
    # 파스텔톤 컬러 팔레트 정의 (겹치지 않는 충분한 수의 색상)
    pastel_colors = [
        '#FFB3B3', '#BAFFC9', '#BAE1FF', '#FFE4BA',  # 파스텔 레드, 그린, 블루, 오렌지
        '#E0BBE4', '#957DAD', '#FEC8D8', '#FFDFD3',  # 파스텔 퍼플, 라벤더, 핑크, 피치
        '#D4F0F0', '#CCE2CB', '#B6CFB6', '#97C1A9',  # 파스텔 민트, 세이지, 그린
        '#FCB9AA', '#FFDBCC', '#ECEAE4', '#A2E1DB',  # 파스텔 코랄, 살몬, 베이지, 터콰이즈
        '#CCD1FF', '#B5EAD7', '#E2F0CB', '#FFDAC1'   # 파스텔 퍼플블루, 민트, 라임, 피치
    ]

    # 각 차트별로 서로 다른 구간의 컬러 사용
    product_colors = pastel_colors[0:8]      # 처음 8개 색상
    material_colors = pastel_colors[8:16]     # 다음 8개 색상
    price_color = pastel_colors[16]          # 17째 색상
    discount_color = pastel_colors[17]       # 18번째 색상

    # Color 매핑 함수 추가
    def get_color(label):
        color_map = {
            'Red': '#FF0000', 'Blue': '#0000FF', 'Green': '#00FF00',
            'Yellow': '#FFFF00', 'Purple': '#800080', 'Orange': '#FFA500',
            'Pink': '#FFC0CB', 'Brown': '#A52A2A', 'Black': '#2C2C2C',
            'White': '#F0F0F0', 'Gray': '#808080', 'Multicolor': '#E5E5E5'
        }
        return color_map.get(label, '#CCCCCC')

    # Common chart layout
    chart_layout = dict(
        height=368,
        margin=dict(t=0, l=30, r=30, b=50),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=12),
            tracegroupgap=5
        ),
        showlegend=True
    )
    
    # Product Assortment Analysis
    category_counts = df['Category'].apply(lambda x: x.split('>')[-1].strip()).value_counts()
    category_counts = category_counts[~category_counts.index.astype(str).str.isnumeric()]
    
    # 4% 미만 필터링을 위한 전처리
    total_products = category_counts.sum()
    category_percentages = (category_counts / total_products) * 100
    filtered_categories = category_percentages[category_percentages > 4]  # 4% 초과만 포함
    
    # 필터링된 데이터로 데이터프레임 생성
    product_dist = pd.DataFrame({
        'Category': filtered_categories.index,
        'Count': category_counts[filtered_categories.index]
    })
    
    # 긴 카테고리명에 줄바꿈 추가
    product_dist['Category'] = product_dist['Category'].apply(
        lambda x: '<br>'.join(textwrap.wrap(x, width=20)) if len(x) > 20 else x
    )
    
    fig_product = px.pie(
        product_dist,
        values='Count',
        names='Category',
        title=None,
        color_discrete_sequence=product_colors
    )
    fig_product.update_layout(**chart_layout)
    fig_product.update_traces(
        textposition='inside',
        textinfo='percent',
        hole=0.3,
        marker=dict(
            line=dict(color='#E5E5E5', width=1)
        )
    )
    
    # Material Analysis
    def extract_materials(materials_str):
        if pd.isna(materials_str):
            return []
        materials = []
        for material in str(materials_str).split(','):
            mat = material.strip().split(' ')[0]
            materials.append(mat)
        return materials

    materials_list = df['Materials'].apply(extract_materials).explode()
    materials_counts = materials_list.value_counts()
    
    # 4% 미만 필터링을 위한 전처리
    total_materials = materials_counts.sum()
    material_percentages = (materials_counts / total_materials) * 100
    filtered_materials = material_percentages[material_percentages > 4]  # 4% 초과만 포함
    
    # 필터링된 데이터로 데이터프레임 생성
    materials_dist = pd.DataFrame({
        'Material': filtered_materials.index,
        'Count': materials_counts[filtered_materials.index]
    })
    
    fig_materials = px.pie(
        materials_dist,
        values='Count',
        names='Material',
        title=None,
        color_discrete_sequence=material_colors
    )
    fig_materials.update_layout(**chart_layout)
    fig_materials.update_traces(
        textposition='inside',
        textinfo='percent',
        hole=0.3,
        marker=dict(
            line=dict(color='#E5E5E5', width=1)
        )
    )
    
    # Color Analysis - 수정된 부분
    if 'Color' in df.columns:
        color_counts = df['Color'].value_counts()
        
        # color_mapping에서 실제 색상값 가져오기
        colors = [color_mapping.get(color, '#CCCCCC') for color in color_counts.index]
        
        fig_colors = go.Figure(data=[go.Pie(
            labels=color_counts.index,
            values=color_counts.values,
            hole=0.3,
            marker=dict(
                colors=colors,  # 실제 매핑된 색상 사용
                line=dict(color='#E5E5E5', width=1)
            ),
            textinfo='percent',
            textposition='inside'
        )])
        
        # 텍스트 색상 자동 조정
        def get_text_color(background_color):
            """텍스트 색상 자동 조정 함수 수정"""
            # 연한 회색 배경인 경우에도 검정색 텍스트 사용
            if background_color in ['#F2F2F2', '#F5F5F5']:
                return '#000000'
            # RGB 값 추출 및 밝기 계산
            r = int(background_color[1:3], 16)
            g = int(background_color[3:5], 16)
            b = int(background_color[5:7], 16)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return '#000000' if luminance > 0.5 else '#FFFFFF'
        
        text_colors = [get_text_color(color) for color in colors]
        fig_colors.update_traces(
            textfont=dict(color=text_colors)
        )
        
        fig_colors.update_layout(
            height=320,
            margin=dict(t=0, l=30, r=30, b=30),
            legend=dict(
                yanchor="middle",
                y=0.5,
                xanchor="right",
                x=1.1,
                itemsizing='constant',
                font=dict(size=13),
                tracegroupgap=10,
                itemwidth=40
            ),
            showlegend=True
        )

    # Price Analysis - Separate charts
    categories = df['Category'].apply(lambda x: x.split('>')[-1].strip()).unique()
    
    # 카테고리가 10개 이상일 때 줄바꿈 적용하는 함수
    def wrap_category_labels(categories):
        if len(categories) > 10:
            return ['\n'.join(textwrap.wrap(cat, width=15)) for cat in categories]
        return categories
    
    # 줄바꿈이 적용된 카테고리 레이블 생성
    wrapped_categories = wrap_category_labels(categories)
    
    # Average Price chart
    fig_prices = go.Figure()
    fig_prices.add_trace(go.Bar(
        x=wrapped_categories,
        y=[round(df[df['Category'].str.contains(cat)]['Original_Price'].mean(), 1) for cat in categories],
        marker_color='rgb(135, 206, 235)'
    ))
    
    # 카테고리 개수에 따라 차트 높이 동적 조정
    chart_height = 300 if len(categories) <= 10 else 400
    
    fig_prices.update_layout(
        title="Average Price per Category (USD)",
        height=chart_height,
        margin=dict(t=30, l=30, r=30, b=100),  # 하단 여백 증가
        yaxis=dict(
            title="Price ($)",
            gridcolor='lightgray',
            zerolinecolor='lightgray'
        ),
        xaxis=dict(
            title="Category",
            tickangle=45 if len(categories) > 10 else 0,  # 카테고리 개수에 따라 각도 조정
            tickmode='array',
            ticktext=wrapped_categories,
            tickvals=list(range(len(categories)))
        ),
        plot_bgcolor='white',
        showlegend=False
    )
    
    # Discount Rate chart
    fig_discounts = go.Figure()
    
    discount_data = []
    for cat in categories:
        category_data = df[df['Category'].str.contains(cat)]
        if not category_data.empty:
            avg_discount = category_data['Discount'].mean()
            discount_data.append({
                'Category': cat,
                'Discount': avg_discount
            })
    
    if discount_data:
        discount_df = pd.DataFrame(discount_data)
        discount_df['Discount'] = discount_df['Discount'].round(1)
        
        fig_discounts.add_trace(go.Bar(
            x=wrapped_categories,
            y=discount_df['Discount'],
            marker_color='rgb(255, 160, 122)'
        ))
        
        fig_discounts.update_layout(
            title="Discount Percentages by Category",
            height=chart_height,
            margin=dict(t=30, l=30, r=30, b=100),  # 하단 여백 증가
            yaxis=dict(
                title="Discount (%)",
                gridcolor='lightgray',
                zerolinecolor='lightgray',
                range=[0, 100],
                tickmode='linear',
                tick0=0,
                dtick=20,
                ticksuffix='%'
            ),
            xaxis=dict(
                title="Category",
                tickangle=45 if len(categories) > 10 else 0,  # 카테고리 개수에 따라 각도 조정
                tickmode='array',
                ticktext=wrapped_categories,
                tickvals=list(range(len(categories)))
            ),
            plot_bgcolor='white',
            showlegend=False
        )

    return fig_product, fig_materials, fig_prices, fig_discounts

def convert_excel_to_df(uploaded_file):
    """업로드된 파일에서 필요한 6개 컬럼만 추출"""
    try:
        # 필요한 컬럼 정의
        required_columns = {
            'Title': str,
            'Category': str,
            'Original Price (USD)': float,
            'Current Price (USD)': float,
            'Materials': str,
            'Discount': float
        }
        
        # 헤더가 4행에 있으므로, 3을 사용 (0-based index)
        df = pd.read_excel(uploaded_file, header=3)
        
        # 내부용 텍스트 포함된 행 찾기 및 제거
        internal_text = "This report is for internal use only. Please refer to our Terms of Service for full details."
        mask = df.apply(lambda x: x.astype(str).str.contains(internal_text, case=False, na=False))
        if mask.any().any():
            # 당 텍스트가 포함된 행의 인덱스 찾기
            internal_row_idx = mask.any(axis=1).idxmax()
            # 해당 행까지의 데이터만 유지
            df = df.iloc[:internal_row_idx]
            
        # 필요한 컬럼만 선택
        try:
            selected_df = df[required_columns.keys()]
            
            # 데이터 타입 변환
            for col, dtype in required_columns.items():
                if dtype == float:
                    selected_df[col] = pd.to_numeric(selected_df[col], errors='coerce')
                else:
                    selected_df[col] = selected_df[col].astype(str)
            
            # 모든 값이 null인 행 제거
            selected_df = selected_df.dropna(how='all')
            
            return selected_df
            
        except KeyError as e:
            st.error(f"필요한 컬럼을 찾을 수 없습니다: {str(e)}")
            st.write("현재 파일의 컬럼:", df.columns.tolist())
            return None
            
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {str(e)}")
        return None

def get_image_hash(image):
    """이미지 해시 """
    return hash(image.tobytes())

def encode_image(image):
    """PIL Image base64로 인코딩"""
    buffered = io.BytesIO()
    # RGBA 모드를 RGB로 변환 후 저장
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    image.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode()

# 이미지 분석 옵션 정의
analysis_options = {
    "top": {
        "category": ["T-Shirt", "Shirt", "Blouse", "Sweater", "Hoodie", "Tank Top", "Cardigan", "Jacket", "Coat"],
        "fit": ["Oversized", "Loose Fit", "Regular Fit", "Slim Fit", "Cropped", "Boxy", "Form-Fitting"],
        "neckline": ["Crew Neck", "V-neck", "Scoop Neck", "Turtleneck", "Mock Neck", "Boat Neck", "Cowl Neck", "Halter Neck", "Square Neck"],
        "sleeves": ["Sleeveless", "Cap Sleeves", "Short Sleeves", "Three Quarter Sleeves", "Long Sleeves", "Dolman Sleeves", "Raglan Sleeves", "Bell Sleeves"],
        "details": ["Basic", "Ruffle", "Pleated", "Gathered", "Color Block", "Embroidered", "Sequined", "Lace Trim", "Button-Down"],
        "pattern": ["Solid", "Striped", "Floral", "Check", "Polka Dot", "Animal Print", "Geometric", "Abstract", "Graphic"]
    },
    "bottom": {
        "category": ["Pants", "Jeans", "Skirt", "Shorts", "Leggings", "Culottes", "Cargo Pants", "Wide Leg Pants"],
        "fit": ["Skinny", "Straight", "Wide Leg", "Bootcut", "Relaxed", "Tapered", "Baggy", "Regular"],
        "length": ["Mini", "Knee Length", "Midi", "Maxi", "Ankle Length", "Full Length", "Cropped", "Three Quarter"],
        "rise": ["Low Rise", "Mid Rise", "High Rise", "Ultra High Rise", "Natural Waist"],
        "details": ["Basic", "Distressed", "Pleated", "Cargo Pockets", "Side Slit", "Raw Hem", "Frayed", "Zip Detail"],
        "pattern": ["Solid", "Striped", "Check", "Herringbone", "Camo", "Animal Print", "Geometric", "Floral"]
    },
    "dress": {
        "category": ["Mini Dress", "Midi Dress", "Maxi Dress", "Shift Dress", "Wrap Dress", "A-Line Dress", "Bodycon Dress", "Slip Dress"],
        "fit": ["Loose", "Regular", "Fitted", "A-Line", "Empire", "Sheath", "Mermaid", "Tent"],
        "neckline": ["V-neck", "Round Neck", "Square Neck", "Sweetheart", "Halter", "Off-Shoulder", "One-Shoulder", "Cowl Neck"],
        "sleeves": ["Sleeveless", "Cap Sleeves", "Short Sleeves", "Three Quarter", "Long Sleeves", "Flutter Sleeves", "Bishop Sleeves", "Puff Sleeves"],
        "details": ["Basic", "Ruched", "Pleated", "Draped", "Tiered", "Ruffled", "Belted", "Wrap Style", "Button Front"],
        "pattern": ["Solid", "Floral", "Polka Dot", "Striped", "Abstract", "Geometric", "Animal Print", "Check"]
    },
    "common": {
        "fabric_type": ["Cotton", "Polyester", "Wool", "Silk", "Linen", "Denim", "Jersey", "Knit", "Leather", "Velvet"],
        "texture": ["Smooth", "Textured", "Ribbed", "Quilted", "Brushed", "Woven", "Mesh", "Lace"],
        "season": ["Spring", "Summer", "Fall", "Winter", "All Season"],
        "style": ["Casual", "Formal", "Business", "Athletic", "Bohemian", "Minimalist", "Streetwear", "Vintage"]
    }
}

def extract_images_from_excel(uploaded_file):
    """엑셀 파일에서 이미지 추출"""
    try:
        file_content = io.BytesIO(uploaded_file.getvalue())
        wb = openpyxl.load_workbook(file_content)
        sheet = wb.active
        image_loader = SheetImageLoader(sheet)
        
        images = []
        processed_cells = set()  # 중복 처리 방지
        
        # 이미지 50개 제한을 위한 카운터
        image_count = 0
        MAX_IMAGES = 50
        
        for row in sheet.iter_rows():
            if image_count >= MAX_IMAGES:
                break
                
            for cell in row:
                try:
                    coord = cell.coordinate
                    if coord not in processed_cells and image_loader.image_in(coord):
                        image = image_loader.get(coord)
                        if image:
                            images.append(image)
                            processed_cells.add(coord)
                            image_count += 1
                            
                            if image_count >= MAX_IMAGES:
                                break
                except Exception as e:
                    continue
        
        # 크북 닫기
        wb.close()
        
        # 첫 번째 이미지(배너) 제외하고 반환
        images = images[1:] if len(images) > 1 else []
        
        # 이미지 수가 50개를 초과하는 경우 경고 메시지 표시
        if len(images) >= MAX_IMAGES:
            st.warning(f"이미지가 {MAX_IMAGES}개를 초과하여, 처음 {MAX_IMAGES}개의 이미지만 분석됩니다.")
        
        return images
        
    except Exception as e:
        st.error(f"이미지 추출 중 오류 발생: {str(e)}")
        return []

def analyze_single_image(image):
    """단일 이미지 분석"""
    try:
        # 이미지를 base64로 인코딩
        base64_image = encode_image(image)
        
        prompt = f"""
        이미지를 보고 해당 의류가 'Top', 'Bottom', 'Dress' 중 어떤 것인지 분류하고, 이에 따라 다음 항목들을 분석해주세요:

        만약 'Top'인 경우:
        1. Fit (착용감): Loose Fit, Regular Fit, Slim Fit 중 선택
        2. Neckline (넥라인): "Crew Neck", "V-Neck", "Square Neck", "Scoop Neck", "Henley Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Sweetheart", "Polo Collar", "Shirt Collar" 중 선택
        3. Sleeves (소매): "Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Half Sleeves", "Puff Sleeves" 중 선택
        4. Color Group (색상 그룹): Neutrals, Vibrant, Pastels, Pattern/Graphic, Earth Tones 중 선택
        5. Pattern (패턴): 패턴이 있다면 리스트의 옵션 중 선택하고 "Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine" 종류 명시, 없으면 Unspecified
        6. Details (디테일): 옵션 중 해당되는 Detail 이 있다면 "Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Appliqué", "Shirring", "Wrap", "Twist", "Knot", "Mix media", "Seam detail", "Cut out", "Seamless", "Binding" 이 중 선택하고 해당하지 않는다면, Unspecified

        만약 'Bottom'인 경우:
        1. Fit (착용감): "Slim Fit", "Regular Fit", "Loose Fit", "Skinny", "Straight", "Bootcut", "Flare", "Wide Leg" 중 선택
        2. Length (길이): "Short", "Knee Length", "Ankle Length", "Full Length" 중 선택
        3. Rise (허리 높이): "Low Rise", "Mid Rise", "High Rise" 중 선택
        4. Color Group (색상 그룹): Neutrals, Vibrant, Pastels, Pattern/Graphic, Earth Tones 중 선택
        5. Pattern (패턴): 패턴이 있다면 리스트의 옵션 중 선택하고 "Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine" 종류 명시, 없으면 Unspecified
        6. Details (디테일): "Distressed", "Ripped", "Embroidery", "Pockets", "Belt Loops", "Pleats" 중 선택, 해당하지 않는다면 Unspecified

        만약 'Dress'인 경우:
        1. Fit (착용감): "Bodycon", "A-Line", "Fit&Flare", "Shift", "Sheath", "Empire Waist" 중 선택
        2. Neckline (넥라인): "Crew Neck", "V-Neck", "Square Neck", "Scoop Neck", "Henley Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Sweetheart", "Polo Collar", "Shirt Collar" 중 선택
        3. Sleeves (소매): "Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Half Sleeves", "Puff Sleeves" 중 선택
        4. Color Group (색상 그룹): Neutrals, Vibrant, Pastels, Pattern/Graphic, Earth Tones 중 선택
        5. Pattern (패턴): 있다면 패턴 종류 명시, 없으면 Unspecified
        6. Details (디테일): 옵션 중 해당되는 Detail 이 있다면 "Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Appliqué", "Shirring", "Wrap", "Twist", "Knot", "Mix media", "Seam detail", "Cut out", "Seamless", "Binding" 이 중 선택하고 해당하지 않는다면, Unspecified

        응답은 JSON 형식으로 해주시고, 'Category' 필드에 'Top', 'Bottom', 'Dress' 중 하나를 포함해주세요.
        """
        
        # GPT-4 API 호출
        response = client.chat.completions.create(
            model="gpt-4o",  # 비전 모델로 변경
            messages=[
                {
                    "role": "system",
                    "content": "You are a fashion design expert. Analyze the image and categorize it appropriately."
                },
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=300
        )
        
        # 응답 파싱
        result = response.choices[0].message.content
        result = result.replace('```json', '').replace('```', '').strip()
        
        try:
            return json.loads(result)
        except json.JSONDecodeError as e:
            st.error(f"JSON 파싱 오류: {str(e)}")
            return None
                
    except Exception as e:
        st.error(f"이미지 분석 중 오류 발생: {str(e)}")
        return None

def get_design_summary(images, analysis_results):
    """이미지들의 디자인 요소를 분석하여 요약"""
    try:
        # 이미지들을 base64로 인코딩
        base64_images = [encode_image(img) for img in images[:3]]  # 처음 3개 이미지만 분석
        
        prompt = f"""
        As a fashion design expert, provide a comprehensive brand and design mood analysis focusing on:

        1. Overall brand identity and aesthetic direction
        2. Key design elements and signature details
        3. Target market positioning and lifestyle appeal
        4. Design philosophy and creative approach
        5. Market trends alignment and uniqueness

        Consider these analysis results:
        {json.dumps(analysis_results, indent=2, ensure_ascii=False)}

        Provide a concise 2-3 sentence summary in English that captures the brand's essence and design direction.
        Focus on the overall mood and brand positioning rather than specific technical details.
        """

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": "You are a fashion design director. Provide practical and concrete design insights based on current trends and data."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            max_tokens=300
        )

        return response.choices[0].message.content

    except Exception as e:
        st.error(f"디자인 요약 생성 중 오류 발생: {str(e)}")
        return "디자인 요약을 생성할 수 없습니다."

def display_image_analytics(images, analysis_results):
    """이미지 분석 결과 표시 UI"""
    
    st.markdown('<p class="design-overview">Design Overview</p>', unsafe_allow_html=True)
    
    # 이미지 개수 표시 제거
    
    # 이미지 표시
    cols = st.columns(3)
    for idx, img in enumerate(images[:3]):
        cols[idx].image(img, use_column_width=True)
    
    # AI 분석 내용
    with st.spinner("Analyzing designs..."):
        design_summary = get_design_summary(images, analysis_results)
        st.markdown(f"""
        📊 **Analytics summary**  
        {design_summary}
        """)
    
    # Analysis Results 섹션
    st.markdown("""
        <style>
        .section-divider {
            border-top: 1px solid #cccccc;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.write("### Analysis Results")
    
    # 복종별 탭 생성
    top_tab, bottom_tab, dress_tab = st.tabs(["Top", "Bottom", "Dress"])
    
    # 파스텔톤 컬러 팔레트
    pastel_colors = [
        '#FFB3B3', '#BAFFC9', '#BAE1FF', '#FFE4BA',
        '#E0BBE4', '#957DAD', '#FEC8D8', '#FFDFD3',
        '#D4F0F0', '#CCE2CB', '#B6CFB6', '#97C1A9',
        '#FCB9AA', '#FFDBCC', '#ECEAE4', '#A2E1DB',
        '#CCD1FF', '#B5EAD7', '#E2F0CB', '#FFDAC1'
    ]
    
    # 복종별 데이터 분리 및 표시
    category_data = {'Top': {}, 'Bottom': {}, 'Dress': {}}
    
    for category in ['Top', 'Bottom', 'Dress']:
        for metric in analysis_results[category]:
            category_data[category][metric] = analysis_results[category][metric]
    
    # 복종별 분석 결과 표시 함수
    def display_category_charts(tab, category_data, category_type):
        with tab:
            # metrics 순서 재정의 - Color Group을 Neckline 다음으로 이동
            if category_type == "Top":
                metrics = ["Fit", "Neckline",  # 첫 번째 줄
                          "Sleeves", "Color Group",  # 두 번째 줄 (Color Group을 여기로 이동)
                          "Pattern", "Details"]  # 세 번째 줄
            elif category_type == "Bottom":
                metrics = ["Fit", "Length",  # 첫 번째 줄
                          "Rise", "Color Group",  # 두 번째 줄 (Color Group을 여기로 이동)
                          "Pattern", "Details"]  # 세 번째 줄
            elif category_type == "Dress":
                metrics = ["Fit", "Neckline",  # 첫 번째 줄
                          "Sleeves", "Color Group",  # 두 번째 줄
                          "Pattern", "Details"]  # 세 번째 줄

            # 2개씩 차 배치
            for i in range(0, len(metrics), 2):
                col1, col2 = st.columns(2)
                
                # 첫 번째 차트
                with col1:
                    metric = metrics[i]
                    if metric in category_data:
                        color_start = (i * 4) % len(pastel_colors)
                        colors = pastel_colors[color_start:color_start+4]
                        
                        fig = px.pie(
                            values=list(category_data[metric].values()),
                            names=list(category_data[metric].keys()),
                            title=metric,
                            hole=0.3,
                            color_discrete_sequence=colors
                        )
                        fig.update_layout(
                            height=360,
                            margin=dict(t=60, l=20, r=20, b=100),
                            showlegend=True,
                            title={
                                'text': metric,
                                'font': {'size': 27},
                                'y': 0.95,
                                'x': 0.5,
                                'xanchor': 'center'
                            },
                            legend=dict(
                                orientation="h",
                                yanchor="top",
                                y=-0.2,
                                xanchor="center",
                                x=0.5,
                                font=dict(size=12)
                            )
                        )
                        fig.update_traces(
                            marker=dict(line=dict(color='#E5E5E5', width=1))
                        )
                        st.plotly_chart(fig, use_container_width=True)
                
                # 두 번째 차트
                if i + 1 < len(metrics):
                    with col2:
                        metric = metrics[i+1]
                        if metric in category_data:
                            color_start = ((i + 1) * 4) % len(pastel_colors)
                            colors = pastel_colors[color_start:color_start+4]
                            
                            fig = px.pie(
                                values=list(category_data[metric].values()),
                                names=list(category_data[metric].keys()),
                                title=metric,
                                hole=0.3,
                                color_discrete_sequence=colors
                            )
                            fig.update_layout(
                                height=360,
                                margin=dict(t=60, l=20, r=20, b=100),
                                showlegend=True,
                                title={
                                    'text': metric,
                                    'font': {'size': 27},
                                    'y': 0.95,
                                    'x': 0.5,
                                    'xanchor': 'center'
                                },
                                legend=dict(
                                    orientation="h",
                                    yanchor="top",
                                    y=-0.2,
                                    xanchor="center",
                                    x=0.5,
                                    font=dict(size=12)
                                )
                            )
                            fig.update_traces(
                                marker=dict(line=dict(color='#E5E5E5', width=1))
                            )
                            st.plotly_chart(fig, use_container_width=True)
    
    display_category_charts(top_tab, category_data['Top'], "Top")
    display_category_charts(bottom_tab, category_data['Bottom'], "Bottom")
    display_category_charts(dress_tab, category_data['Dress'], "Dress")

def analyze_images(images):
    """이미지 분석 및 결과 집계"""
    # 이미지 50개로 제한
    MAX_IMAGES = 50
    if len(images) > MAX_IMAGES:
        images = images[:MAX_IMAGES]
    
    analysis_results = {
        "Top": {},
        "Bottom": {},
        "Dress": {}
    }
    
    # 배치 크기 증가로 처리 속도 개선
    BATCH_SIZE = 20  # 기존 10에서 20으로 증가
    total_images = len(images)
    
    # 프로그레스 바 생성
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 병렬 처리를 위한 이미지 배치 구성
    for i in range(0, total_images, BATCH_SIZE):
        batch = images[i:i + BATCH_SIZE]
        
        # 현재 진행 상황 업데이트
        current_progress = min((i + BATCH_SIZE) / total_images, 1.0)
        progress_bar.progress(current_progress)
        status_text.text(f"Analyzing... {int(current_progress * 100)}%")
        
        # 배치 내 각 이미지 석 - 응답 대기 시간 최적화
        for image in batch:
            try:
                result = analyze_single_image(image)
                if result is None:
                    continue
                    
                # 결과가 문자열인 경우 JSON으로 파싱
                if isinstance(result, str):
                    result = json.loads(result)
                
                # 카테고리 확인 - 대소문자 구분 없이 처리
                category = None
                for key in result:
                    if key.lower() == 'category':
                        category = result[key]
                        break
                
                if not category or category not in analysis_results:
                    continue
                
                # 결과 집계 최적화
                for key, value in result.items():
                    if key.lower() == 'category':
                        continue
                        
                    analysis_results[category].setdefault(key, {})
                    if isinstance(value, str):
                        analysis_results[category][key][value] = analysis_results[category][key].get(value, 0) + 1
                        
            except Exception as e:
                continue
    
    # 분석 완료 메시지
    progress_bar.empty()
    status_text.empty()
    st.success("The image analysis is complete.")
    
    return analysis_results

def get_business_insights(data_summary):
    """영업팀을 위한 비즈니스 인사이트 생성"""
    try:
        prompt = f"""
        패션 산업 전문가로서 다음 데이터를 분석하여 영업팀을 위한 실행 가능한 인사이트를 제공해주세요:

        데이터 요약:
        {json.dumps(data_summary, indent=2, ensure_ascii=False)}

        다음 영역에 대해 구체적인 분석을 제공해주세요:
        1. 시장 기회 및 위험 요소
        2. 가격 전략 및 할인 정책
        3. 제품 포트폴리오 최적화
        4. 시즌별 판매 전략
        5. 구체적인 실행 계획

        응답은 다음 키를 포함한 JSON 형식으로 작성해주세요:
        - market_analysis
        - pricing_strategy
        - portfolio_optimization
        - seasonal_strategy
        - action_plan
        """

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": "You're a fashion apparel vendor sales professional with 15 years of experience. You know all too well what your customers (buyers) want. You translate this into practical and concrete business insights based on data."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.7
        )

        return json.loads(response.choices[0].message.content)

    except Exception as e:
        st.error(f"비즈니스 인사이트 생성 중 오류 발생: {str(e)}")
        return None

def get_design_insights(image_analysis_results):
    """디자인 인사이트 생성"""
    try:
        prompt = f"""
        패션 디자인 전문가로서 다음 이미지 분석 결과를 바탕으로 디자인 인사이트를 제공해주세요:

        석 결과:
        {json.dumps(image_analysis_results, indent=2, ensure_ascii=False)}

        다음 영역에 대해 구체적인 제안을 제공해주세요:
        1. 현재 디자인 트렌드 분석
        2. 개선이 필요한 디자인 요소
        3. 새로운 디자인 제안
        4. 재질 및 컬러 믹스 전략
        5. 다음 시즌 디자인 방향성

        응답은 다음 키를 포함한 JSON 형식으로 작성해주세요:
        - trend_analysis
        - design_improvements
        - new_design_suggestions
        - material_color_strategy
        - next_season_direction
        """

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": "당신은 패션 디자인 디렉터입니다. 현재 트렌드와 분석 데이터를 바탕으로 실용적인 디자 인사이트를 제공합니다."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.7
        )

        return json.loads(response.choices[0].message.content)

    except Exception as e:
        st.error(f"디자인 인사이트 생성 중 오류 발생: {str(e)}")
        return None

def export_to_pdf(df, figures, analysis_results, insights):
    """분석 결과를 PDF로 내보내기"""
    try:
        # PDF 생성을 위한 메모리 버퍼
        buffer = BytesIO()
        
        # PDF 문서 생성
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # 스타일 설정
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            spaceAfter=30
        )
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=18,
            spaceAfter=20
        )
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=12
        )
        
        # PDF에 들어갈 요소들
        elements = []
        
        # 제목
        elements.append(Paragraph("White Space Analysis Report", title_style))
        elements.append(Spacer(1, 20))
        
        # 주요 지표
        elements.append(Paragraph("Key Metrics", heading_style))
        metrics_text = f"""
        Total SKUs: {len(df)}<br/>
        Average Price: ${df['Current_Price'].mean():.2f}<br/>
        Average Discount: {df['Discount'].mean():.1f}%
        """
        elements.append(Paragraph(metrics_text, body_style))
        elements.append(Spacer(1, 20))
        
        # 차트 추가
        for fig, title in figures:
            # Plotly 차트를 HTML로 변환
            html_str = fig.to_html(include_plotlyjs=True, full_html=False)
            
            # HTML을 이미지로 변환
            img_data = f"""
            <div style="text-align: center;">
                <h3>{title}</h3>
                {html_str}
            </div>
            """
            elements.append(Paragraph(img_data, body_style))
            elements.append(Spacer(1, 20))
        
        # AI 인사이트 추가
        if insights:
            elements.append(Paragraph("AI-Powered Insights", heading_style))
            elements.append(Paragraph(insights, body_style))
        
        # PDF 생성
        doc.build(elements)
        
        # 버퍼의 내용을 바이트로 가져오기
        pdf_bytes = buffer.getvalue()
        buffer.close()
        
        return pdf_bytes
        
    except Exception as e:
        st.error(f"PDF 생성 중 오류 발생: {str(e)}")
        return None

# Main app
def main():
    st.title("⚪ Yakjin Brand Analyzer")
    
    # 파일 업로드 UI
    uploaded_file = st.file_uploader("Upload CMI Excel File", type=['csv', 'xlsx', 'xls'])
    
    # Overview 탭을 제거하고 2개의 탭만 생성
    tab1, tab2 = st.tabs(["Product Data Analytics", "Image Data Analytics"])
    
    if uploaded_file:
        # 업로더 숨기기
        st.markdown("""
            <style>
                [data-testid="stFileUploader"] {
                    display: none;
                }
            </style>
            """, unsafe_allow_html=True)
        
        # 파일을 데이터프레임으로 변환
        raw_df = convert_excel_to_df(uploaded_file)
        
        if raw_df is not None:
            # 데이터 정제
            df = clean_dataframe(raw_df)
            
            if df is not None and not df.empty:
                # Product Data Analytics 탭 내용
                with tab1:
                    # Display metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total SKUs", df.shape[0])
                    with col2:
                        avg_price = df['Current_Price'].mean()
                        if pd.notna(avg_price):
                            st.metric("Average Price", f"${int(avg_price)}")
                        else:
                            st.metric("Average Price", "N/A")
                    with col3:
                        avg_discount = df['Discount'].mean()
                        if pd.notna(avg_discount):
                            st.metric("Average Discount", f"{int(avg_discount)}%")
                        else:
                            st.metric("Average Discount", "N/A")
                    with col4:
                        st.metric("Categories", df['Category'].nunique())
                    
                    # 프로그레스 바 추가
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # 차트 생성 및 표시
                    fig_product, fig_materials, fig_prices, fig_discounts = analyze_data(df, uploaded_file)
                    
                    # Product Assortment Analysis
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<p class="chart-title">Product Assortment Analysis</p>', unsafe_allow_html=True)
                    st.plotly_chart(fig_product, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 프로그레스 바 업데이트 (33%)
                    progress_bar.progress(33)
                    status_text.text("Generating insights... 33%")
                    
                    # Material Analysis
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<p class="chart-title">Material Composition Analysis</p>', unsafe_allow_html=True)
                    st.plotly_chart(fig_materials, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 프로그레스 바 업데이트 (66%)
                    progress_bar.progress(66)
                    status_text.text("Generating insights... 66%")
                    
                    # Price & Discount Analysis
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<p class="chart-title">Price & Discount Analysis</p>', unsafe_allow_html=True)
                    st.plotly_chart(fig_prices, use_container_width=True)
                    st.plotly_chart(fig_discounts, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # AI Insights
                    data_summary = prepare_data_summary(df)
                    insights = get_ai_insights(data_summary)
                    if insights:
                        st.write("---")
                        st.markdown(insights)
                    
                    # 프로그레스 바 완료 및 제거
                    progress_bar.progress(100)
                    status_text.text("Generating insights... 100%")
                    time.sleep(0.5)  # 완료 상태를 잠시 보여줌
                    progress_bar.empty()
                    status_text.empty()
                
                # Image Data Analytics 탭 내용
                with tab2:
                    st.subheader("Image Data Analytics")
                    
                    # 이미지 추출
                    with st.spinner("이미지를 분석하고 있습니다..."):
                        images = extract_images_from_excel(uploaded_file)
                        
                        if not images:
                            st.warning("분석 가능한 제품 이미지를 찾을 수 없습니다.")
                            return
                        
                        # 이미지 분석 및 결과 표시
                        analysis_results = analyze_images(images)
                        display_image_analytics(images, analysis_results)               
    else:
        # 파일이 업로드되지 않은 경우 메시지 표시 (tab1으로 변경)
        with tab1:
            st.info("👆 분석을 시작하려면 CMI 데이터 파일을 업로드해주세요.")

if __name__ == "__main__":
    main()
