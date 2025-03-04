import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
import matplotlib.pyplot as plt
import numpy as np

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Border, Side  # 必要なモジュールをインポート


import seaborn as sns
from collections import Counter

import re

# Secretsからパスワードを取得
PASSWORD = st.secrets["PASSWORD"]

# パスワード認証の処理
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "login_attempts" not in st.session_state:
    st.session_state.login_attempts = 0

def verificar_contraseña():
    contraseña_ingresada = st.text_input("Introduce la contraseña:", type="password")

    if st.button("Iniciar sesión"):
        if st.session_state.login_attempts >= 3:
            st.error("Has superado el número máximo de intentos. Acceso bloqueado.")
        elif contraseña_ingresada == PASSWORD:  # Secretsから取得したパスワードで認証
            st.session_state.authenticated = True
            st.success("¡Autenticación exitosa! Marque otra vez el botón 'Iniciar sesión'.")
        else:
            st.session_state.login_attempts += 1
            intentos_restantes = 3 - st.session_state.login_attempts
            st.error(f"Contraseña incorrecta. Te quedan {intentos_restantes} intento(s).")
        
        if st.session_state.login_attempts >= 3:
            st.error("Acceso bloqueado. Intenta más tarde.")

if st.session_state.authenticated:
    # 認証成功後に表示されるメインコンテンツ


    # 設定: アプリケーションの基本設定
    st.title("Simulación de Proyecto de Inversión por PyME")
    st.write("###### :blue[Esta herramienta facilita una simulación sencilla de inversión por PyMEs para maquinaria o nuevo negocio. Es recomendable realizar los estudios más detallados, al concretar el plan de inversión.]") 
    st.write("###### Ingrese los datos principales del proyecto de inversión a analizar.")

    # 初期値の設定
    # Col1, Col2, Col3の作成（間にスペースを挟む）
    col1, col2_space, col2, col3_space, col3 = st.columns([0.9, 0.05, 0.8, 0.05, 1.25])
    with col1: 
        st.write("###### :red[Maquinaria o Equipo a comprar (O, inversión para el nuevo negocio):]")
        inversion_inicial = st.number_input("Monto de inversión (Lps)", value=100000)
        vida_util = st.number_input("Años de vida útil del equipo a invertir", value=6)
        st.write("###### :red[Tasa de impuesto:]")
        tasa_impuesto = st.number_input("Tasa de impuesto (%)", value=15)

    with col2:
        st.write("###### :red[Posible Uso de Crédito:]")
        monto_prestamo = st.number_input("Monto del préstamo a aplicar para la inversión (Lps)", value=60000)
        tasa_interes = st.number_input("Tasa de interés del préstamo (%)", value=20)
        meses_prestamo = st.number_input("Plazo del préstamo (meses)", value=30)
        
    with col3:
        st.write("###### :red[Ingresos y costos del proyecto:]")
        ventas_anuales = st.number_input("Ventas anuales (adicionales) a generar por el proyecto (Lps)", value=90000)
        costos_ventas = st.number_input("Proporción (%) de costos productivos sobre las ventas (Nota: Los costos productivos son de materias primas, trabajadores productivos, y otros relacionados al proyecto, excluyendo depreciación)", value=60)
        gastos_administrativos = st.number_input("Gastos administrativos anuales relacionado al proyecto (Lps)", value=5000)

    # Analizarボタンの設定
    if st.button("Analizar"):

        # 償還表作成
        # 月利の計算
        monthly_rate = tasa_interes / 100 / 12
        
        # 月数に基づいて毎月の返済額の計算
        monthly_payment = (monto_prestamo * monthly_rate * (1 + monthly_rate) ** meses_prestamo) / ((1 + monthly_rate) ** meses_prestamo - 1)

        # 初期設定
        balance = monto_prestamo
        schedule = []

        # 各月の償還表を作成
        for month in range(1, meses_prestamo + 1):
            interest_payment = balance * monthly_rate
            principal_payment = monthly_payment - interest_payment
            balance -= principal_payment
            schedule.append([month, round(monthly_payment), round(principal_payment), round(interest_payment), round(balance)])

        # データフレームに変換し、インデックスを表示しない
        df = pd.DataFrame(schedule, columns=["Mes", "Pago mensual (Lps)", "Pago a capital (Lps)", "Interés (Lps)", "Saldo restante (Lps)"])

        st.subheader("A) Cuadro de Amortización del crédito a solicitar en base al plan de cuotas niveladas")
        st.dataframe(df.reset_index(drop=True))

        # 年間利息支払額の計算
        df['Año'] = (df['Mes'] - 1) // 12 + 1  # 各行に対応する年を計算
        yearly_interest = df.groupby('Año')['Interés (Lps)'].sum().reset_index()  # 年ごとの利息の合計を計算

        # 年間元本支払額の計算
        yearly_capital = df.groupby('Año')['Pago a capital (Lps)'].sum().reset_index()

        # 月単位の調整
        full_years = meses_prestamo // 12
        remaining_months = meses_prestamo % 12

        # 年ごとの金利負担
        intereses = yearly_interest['Interés (Lps)'].tolist()[:full_years]

        # 端数調整（月単位での追加）
        if remaining_months > 0:
            third_year_interest = df[df['Año'] == full_years + 1]['Interés (Lps)'].sum()
            intereses.append(third_year_interest)

        # 金利負担のない月はゼロ表示
        intereses += [0] * (vida_util - full_years - 1)

        # 予想損益計算書の作成
        st.subheader("B) Estado de Resultados Proyectado")
        ventas = np.array([ventas_anuales] * vida_util)
        costo_ventas_sin_depreciacion = ventas * (costos_ventas / 100)
        depreciacion = inversion_inicial / vida_util
        costo_total_ventas = costo_ventas_sin_depreciacion + depreciacion
        utilidad_bruta = ventas - costo_total_ventas
        utilidad_operativa = utilidad_bruta - gastos_administrativos

        # 確実に配列として扱うために、interesesをNumPy配列に変換
        intereses = np.array(intereses)

        # 利益に関する計算もNumPy配列として扱う
        utilidad_operativa = np.array(utilidad_operativa)

        # 税前利益・純利益の計算
        utilidad_antes_impuestos = utilidad_operativa - intereses  # 配列同士の引き算
        utilidad_neta = utilidad_antes_impuestos * (1 - tasa_impuesto/100)

        # 損益計算書のデータフレーム
        data_sonkei = {
            "Año": list(range(1, vida_util + 1)),
            "Ventas": ventas,
            "Costos productivos": costo_ventas_sin_depreciacion,
            "Depreciación": [depreciacion] * vida_util,
            "Costo total de ventas": costo_total_ventas,
            "Utilidad bruta": utilidad_bruta,
            "Gastos administrativos": [gastos_administrativos] * vida_util,
            "Intereses": intereses,
            "Utilidad antes de impuestos": utilidad_antes_impuestos,
            "Utilidad neta": utilidad_neta,
        }
        df_sonkei = pd.DataFrame(data_sonkei).T.round(0)  # 小数点以下を四捨五入して整数表示
        st.dataframe(df_sonkei)

        st.write("Nota: Si la utilidad antes de impuestos es negativa, la utilidad neta también debería mostrar una cantidad negativa equivalente. Sin embargo, dado que otros proyectos de la misma empresa podrían generar ganancias, en este cuadro la utilidad neta siempre se presenta como Utilidad antes de impuestos X (1-tasa de impuesto).")

        # キャッシュフロー計算書
        st.subheader("C) Estado de Flujo de Caja Proyectado")

        # flujo_operativoの定義（vida_utilの年数に合わせる）現在はゼロ
        flujo_operativo = [0] + list(utilidad_neta + depreciacion)

        # flujo_inversionの定義（vida_utilに合わせる、最初の年に-inversion_inicial、それ以外は0）
        flujo_inversion = [-inversion_inicial] + [0] * vida_util

        # flujo_financieroの定義（年数をvida_utilに合わせる）
        flujo_financiero = [monto_prestamo] + [-capital for capital in yearly_capital['Pago a capital (Lps)']]
        
        # flujo_operativoがvida_utilに合うように長さを調整
        if len(flujo_operativo) < vida_util + 1:
            flujo_operativo += [0] * (vida_util + 1 - len(flujo_operativo))

        # flujo_inversionの定義
        flujo_inversion = [-inversion_inicial] + [0] * vida_util

        # flujo_financieroの長さをvida_utilに合わせる
        if len(flujo_financiero) < vida_util + 1:
            flujo_financiero += [0] * (vida_util + 1 - len(flujo_financiero))

        # 各リストの長さを確認
        print(f"flujo_operativo: {len(flujo_operativo)}")
        print(f"flujo_inversion: {len(flujo_inversion)}")
        print(f"flujo_financiero: {len(flujo_financiero)}")

        # リストの長さが一致していることを確認
        assert len(flujo_operativo) == len(flujo_inversion) == len(flujo_financiero), "リストの長さが一致していません"

        # flujo_totalの計算
        flujo_total = [flujo_operativo[i] + flujo_inversion[i] + flujo_financiero[i] for i in range(len(flujo_operativo))]

        # キャッシュフローのデータフレーム
        data_cf = {
            "Año": ["Hoy"] + list(range(1, vida_util + 1)),
            "Flujo operativo": flujo_operativo,
            "Flujo de inversión": flujo_inversion,
            "Flujo financiero": flujo_financiero,
            "Flujo neto": flujo_total
        }

        df_cf = pd.DataFrame(data_cf)

        # 数値カラムだけを整数に変換
        numeric_cols = ["Flujo operativo", "Flujo de inversión", "Flujo financiero", "Flujo neto"]
        df_cf[numeric_cols] = df_cf[numeric_cols].round(0).astype(int)

        # データフレームの転置
        df_cf_transposed = df_cf.T

        # 転置後にヘッダー行を設定
        df_cf_transposed.columns = df_cf_transposed.iloc[0]
        df_cf_transposed = df_cf_transposed[1:]

        # データフレームの表示
        st.dataframe(df_cf_transposed)

        # 投資プロジェクト評価指標の作成
        st.subheader("D) Indicadores de Evaluación del Proyecto")
        flujo_operativoOR = list(utilidad_neta + depreciacion)
        flujo_descuento = flujo_operativoOR / ((1 + tasa_interes / 100) ** np.arange(1, vida_util + 1))
        npv = np.sum(flujo_descuento) - inversion_inicial
        roi = np.sum(utilidad_antes_impuestos) / inversion_inicial
        
        rate = tasa_interes/100
        payback = 1/rate - (1/(rate*(1+rate)**vida_util))
        st.write(f"###### Valor Presente Neto (VPN): {npv:.2f} Lps")
        st.write(f"###### Rentabilidad sobre la Inversión (ROI): {roi:.1f} %")
        st.write(f"###### Periodo máximo aceptable para recaudación del fondo invertido: {payback*12:.1f} meses")
    
        st.write("###### :red[Un proyecto con el VPN negativo o insuficiente se debe rechazar. Para simplificar el calculo del VPN, se aplica la tasa de interes, como la tasa de descuentos. El tercer indicador es para la referencia teórica, y el empresario deberá recuperar el fondo invertido lo antes posible. Se presenta abajo una figura del flujo neto de caja del Proyecto.]") 

        # 棒グラフの作成
        fig, ax = plt.subplots()
        ax.bar(range(vida_util + 1), flujo_total, label='Flujo neto', color='blue')

        # X軸に年ごとのラベルを追加
        ax.set_xticks(range(vida_util + 1))
        ax.set_xticklabels([f'Año {i}' for i in range(vida_util + 1)])

        # 金額ゼロのところに水平線を追加
        ax.axhline(0, color='red', linewidth=1.5)

        # グラフのラベルとタイトル
        ax.set_xlabel('Año')
        ax.set_ylabel('Flujo de caja (Lps)')
        ax.set_title('Proyección de Flujo de caja durante el Proyecto')

        # グラフをStreamlitで表示
        st.pyplot(fig)

        # 損益分岐点分析グラフ
        st.subheader("E) Gráfico de Análisis de Punto de Equilibrio al año")
        st.write("Se presentan abajo el resultado del análisis del punto de equilibrio. El resultado del análisis podrá ser impreciso, en los siguientes dos sentidos. `Primero, en esta simulacion, los costos se clasifican en los fijos y variables de manera no precisa. Segundo, este análisis no incluye el cálculo de descuentos basado en la teoría financiera. Es decir, considerando el costo de adquisición del capital, el punto de equilibrio, en términos reales, podrá ser más alto que la cifra indicada abajo.")

        # 固定費と変動費の計算
        fixedcost = gastos_administrativos + np.mean(intereses) + inversion_inicial/vida_util # 固定費は管理費と平均利息を加えたもの
        variable_ratio = costos_ventas/100  # 変動費率を計算

        # 損益分岐点の計算
        breakeven_sales = fixedcost / (1 - variable_ratio)

        # グラフの作成
        fig, ax = plt.subplots()

        # 損益分岐点前後の売上範囲を設定
        sales_range = np.arange(int(breakeven_sales * 0.8), int(breakeven_sales * 1.2), 100)

        # 総コストを計算
        total_costs = [fixedcost + (variable_ratio * s) for s in sales_range]

        # 総コストと売上のプロット
        ax.plot(sales_range, total_costs, color='skyblue', label="Costos totales", marker='o')
        ax.plot(sales_range, sales_range, color='orange', label="Venta anual", marker='o')

        # グラフのタイトルとラベル
        ax.set_title("Estimación del punto de equilibrio")
        ax.set_xlabel("Venta (Lps)")
        ax.set_ylabel("Costos y ventas (Lps)")

        # 損益分岐点の縦線を追加
        ax.axvline(breakeven_sales, color='red', linestyle='--', label=f"Punto de equilibrio: {breakeven_sales:.0f} Lps")

        # 損益分岐点の説明
        ax.fill_between(sales_range, total_costs, sales_range, where=[s > breakeven_sales for s in sales_range], color='skyblue', alpha=0.3, interpolate=True)

        # グラフに説明を追加
        mid_x = breakeven_sales * 1.05  # 説明テキストの位置調整
        mid_y = (max(total_costs) + max(sales_range)) / 2
        ax.text(mid_x, mid_y, "Ganancia = Área del color azul claro", color="blue", fontsize=7, ha="left")

        # 凡例の表示
        ax.legend()

        # グラフをStreamlitに表示
        st.pyplot(fig)

else:
    verificar_contraseña()
