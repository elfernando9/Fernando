import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, html, dcc
import os

# Especificamos la ruta del archivo Excel
archivo_path = r'C:\Users\Gpon Networks\Desktop\DASHGPON.xlsx'

# Función para procesar datos
def procesar_datos():
    if os.path.exists(archivo_path):
        try:
            # Leer la hoja 'HOY'
            df = pd.read_excel(archivo_path, sheet_name='HOY')
            
            # Convertir la columna 'FECHA_DE_INSTALACIÓN' a tipo datetime
            df['FECHA_DE_INSTALACIÓN'] = pd.to_datetime(df['FECHA_DE_INSTALACIÓN'], errors='coerce')
            
            # Filtrar por SEGUIMIENTO = "Instalado" y año 2024
            df_instalados = df[(df['SEGUIMIENTO'] == 'Instalado') & (df['FECHA_DE_INSTALACIÓN'].dt.year == 2024)].copy()
            
            # Extraer mes y agregar columna
            df_instalados['MES'] = df_instalados['FECHA_DE_INSTALACIÓN'].dt.month_name(locale='es_ES')
            
            # Agrupar por mes y contar instalaciones
            grouped_df = df_instalados.groupby('MES').size().reset_index(name='Cantidad_Instalados')

            # Calcular el total de instalaciones del año
            total_instalados = grouped_df['Cantidad_Instalados'].sum()

            # Ordenar meses en orden cronológico
            ordered_months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                              'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            grouped_df['MES'] = pd.Categorical(grouped_df['MES'], categories=ordered_months, ordered=True)
            grouped_df = grouped_df.sort_values('MES')

            # Calcular el porcentaje de instalaciones por mes
            grouped_df['Porcentaje'] = (grouped_df['Cantidad_Instalados'] / total_instalados) * 100

            return grouped_df, total_instalados
        except Exception as e:
            print(f"Ocurrió un error al procesar el archivo: {e}")
            return pd.DataFrame(), 0
    else:
        print(f"El archivo no fue encontrado en la ruta especificada: {archivo_path}")
        return pd.DataFrame(), 0

# Crear la aplicación Dash
app = Dash(__name__)

# Procesar los datos
grouped_df, total_instalados = procesar_datos()

# Crear gráficas solo si hay datos disponibles
if not grouped_df.empty:
    # Gráfica de líneas
    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=grouped_df['MES'],
        y=grouped_df['Cantidad_Instalados'],
        mode='lines+markers+text',
        text=grouped_df['Cantidad_Instalados'],
        textposition='top center',
        line=dict(color='blue', width=3),
        marker=dict(size=8, color='orange'),
        name=f"Total del año: {total_instalados}"
    ))
    fig_line.update_layout(
        title="<b>Cantidad de Instalaciones por Mes en 2024</b>",
        xaxis=dict(title='Meses', tickangle=45),
        yaxis=dict(title='Cantidad de Instalaciones'),
        height=500,
        width=1000,
        plot_bgcolor='rgba(245,245,245,1)',
        margin=dict(l=60, r=60, t=100, b=60),
    )

    # Gráfica de dona
    fig_donut = px.pie(
        grouped_df,
        names='MES',
        values='Cantidad_Instalados',
        title="Porcentaje de Instalaciones por Mes en 2024",
        hole=0.4,
        color='MES',
        color_discrete_sequence=px.colors.qualitative.Bold
    )
    fig_donut.update_traces(
        textinfo='percent+label',
        textfont=dict(size=16, color="#000000"),
    )
    fig_donut.update_layout(
        title=dict(
            text="<b>Porcentaje de Instalaciones por Mes en 2024</b>",
            font=dict(size=24),
            x=0.5
        ),
        height=500,
        width=1000,
        plot_bgcolor="#f7f7f7",
        margin=dict(l=60, r=60, t=100, b=60),
    )
else:
    fig_line = go.Figure().add_annotation(text="No hay datos disponibles", x=0.5, y=0.5, showarrow=False)
    fig_donut = go.Figure().add_annotation(text="No hay datos disponibles", x=0.5, y=0.5, showarrow=False)

# Definir el diseño de la página
app.layout = html.Div([
    html.H1("Gráficas de Instalaciones 2024", style={'textAlign': 'center'}),
    dcc.Graph(figure=fig_line),
    dcc.Graph(figure=fig_donut)
])

# Ejecutar la aplicación
if __name__ == '__main__':
    app.run_server(debug=True)
